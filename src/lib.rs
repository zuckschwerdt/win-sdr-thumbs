use std::{
    cell::RefCell,
    ffi::OsStr,
    // fs::OpenOptions,
    // io::Write,
    os::windows::prelude::OsStrExt,
    panic::{catch_unwind, AssertUnwindSafe},
    sync::{
        atomic::{
            AtomicPtr, 
            AtomicU32, 
            Ordering
        }, 
        Arc,
        Mutex
    },
};

use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::{
            Direct2D::{
                *, 
                self, 
                Common::*
            },
            Direct3D,
            Direct3D11,
            Dxgi,
            Gdi,
        },
        System::{
            self,
            Com,
            Ole,
            Variant::*,
            Registry::{
                *,
                RegCreateKeyExW,
                RegSetValueExW,
            }
        },
        UI::Shell,
        Data::Xml::MsXml,
        Data::Xml::MsXml::*,
    },
};

// Use correct registry view for 32-bit process on 64-bit Windows
// If read access needed later add KEY_READ to both
#[cfg(target_pointer_width = "32")]
const WRITE_FLAGS: REG_SAM_FLAGS = KEY_WRITE | KEY_WOW64_64KEY;

#[cfg(not(target_pointer_width = "32"))]
const WRITE_FLAGS: REG_SAM_FLAGS = KEY_WRITE;

// =================================================================
//                  FFI Panic Safety Macro
// =================================================================

/// Macro to wrap FFI functions with panic protection.
/// This eliminates the boilerplate code for catch_unwind and error handling.
macro_rules! ffi_guard {
    // For functions that return Result<T>
    (Result<$ret_type:ty>, $body:expr) => {{
        let result = catch_unwind(AssertUnwindSafe(|| $body));
        match result {
            Ok(Ok(value)) => Ok(value),
            Ok(Err(e)) => Err(e),
            Err(_) => {
                //log_message("A PANIC occurred in FFI function.");
                Err(E_FAIL.into())
            }
        }
    }};
    
    // For functions that return HRESULT directly
    (HRESULT, $body:expr) => {{
        let result = catch_unwind(AssertUnwindSafe(|| $body));
        match result {
            Ok(hr) => hr,
            Err(_) => {
                //log_message("A PANIC occurred in FFI function.");
                E_FAIL
            }
        }
    }};
    
    // For functions that return BOOL
    (BOOL, $body:expr) => {{
        let result = catch_unwind(AssertUnwindSafe(|| $body));
        match result {
            Ok(success) => success.into(),
            Err(_) => {
                //log_message("A PANIC occurred in FFI function.");
                false.into()
            }
        }
    }};
}

// RAII wrapper for COM initialization - automatically calls CoUninitialize when dropped
struct ComGuard {
    initialized_by_us: bool,
}

impl ComGuard {
    fn new() -> Result<Self> {
        let hr: HRESULT = unsafe { Com::CoInitializeEx(None, Com::COINIT_APARTMENTTHREADED) };
        match hr {
            S_OK => {
                // We successfully initialized COM, so we're responsible for cleanup
                Ok(Self { initialized_by_us: true })
            }
            S_FALSE => {
                // COM was already initialized by someone else, don't clean up
                Ok(Self { initialized_by_us: false })
            }
            _ => {
                // Real error
                Err(Error::new(hr, "Failed to initialize COM"))
            }
        }
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        if self.initialized_by_us {
            unsafe { Com::CoUninitialize() };
        }
    }
}

// --- Thread-local storage for COM objects that cannot be shared between threads ---
thread_local! {
    static COM_GUARD: RefCell<Option<ComGuard>> = RefCell::new(None);
    static D2D_FACTORY: RefCell<Option<ID2D1Factory1>> = RefCell::new(None);
    static D2D_DEVICE: RefCell<Option<ID2D1Device>> = RefCell::new(None);
    static D2D_CONTEXT: RefCell<Option<ID2D1DeviceContext5>> = RefCell::new(None);
    // This flag tracks if the D2D resources are in a bad state and need to be recreated.
    static D2D_RESOURCES_POISONED: std::cell::Cell<bool> = std::cell::Cell::new(false);
}
/// Initializes and retrieves the thread-local Direct2D and WIC resources.
/// This function ensures that the heavyweight factory and device objects are created only once per thread.
fn get_d2d_resources() -> Result<(ID2D1Factory1, ID2D1Device, ID2D1DeviceContext5)> {
    // If the resources were marked as poisoned (like by a previous EndDraw failure), clear the cached device and context. They will be recreated below.
    if D2D_RESOURCES_POISONED.get() {
        D2D_DEVICE.with(|d| d.borrow_mut().take());
        D2D_CONTEXT.with(|c| c.borrow_mut().take());
        // Reset the flag now that we've cleared the caches.
        D2D_RESOURCES_POISONED.set(false);
    }

    // Get or create the Direct2D Factory.
    let d2d_factory = D2D_FACTORY.with(|factory| -> Result<ID2D1Factory1> {
        let mut factory_ref = factory.borrow_mut();
        if factory_ref.is_none() {
            // Ensure COM is initialized with proper cleanup tracking
            COM_GUARD.with(|guard| -> Result<()> {
                let mut guard_ref = guard.borrow_mut();
                if guard_ref.is_none() {
                    *guard_ref = Some(ComGuard::new()?);
                }
                Ok(())
            })?;
            
            let options = D2D1_FACTORY_OPTIONS {
                debugLevel: D2D1_DEBUG_LEVEL_NONE,
            };
            let d2d: ID2D1Factory1 = unsafe { D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, Some(&options))? };
            *factory_ref = Some(d2d);
        }
        Ok(factory_ref.as_ref().unwrap().clone())
    })?;

    // Get or create the Direct2D Device. This requires a backing D3D11 device.
    let d2d_device = D2D_DEVICE.with(|device| -> Result<ID2D1Device> {
        let mut device_ref = device.borrow_mut();
        if device_ref.is_none() {
            // 1. Create the D3D11 Device
            let mut d3d_device: Option<Direct3D11::ID3D11Device> = None;
            unsafe {
                Direct3D11::D3D11CreateDevice(
                    None,
                    Direct3D::D3D_DRIVER_TYPE_HARDWARE,
                    HMODULE::default(),
                    Direct3D11::D3D11_CREATE_DEVICE_BGRA_SUPPORT, // Required for D2D interop
                    None,
                    Direct3D11::D3D11_SDK_VERSION,
                    Some(&mut d3d_device),
                    None,
                    None,
                )?;
            }
            let dxgi_device: Dxgi::IDXGIDevice = d3d_device.ok_or_else(|| Error::new(E_FAIL, "Failed to create D3D11 device"))?.cast()?;

            // 2. Create the D2D Device from the D3D11 device
            let d2d_dev: ID2D1Device = unsafe { d2d_factory.CreateDevice(&dxgi_device)? };
            *device_ref = Some(d2d_dev);
        }
        Ok(device_ref.as_ref().unwrap().clone())
    })?;

    // Get or create the Direct2D Device Context (expensive, so cache it)
    let d2d_context = D2D_CONTEXT.with(|context| -> Result<ID2D1DeviceContext5> {
        let mut context_ref = context.borrow_mut();
        if context_ref.is_none() {
            let dc: ID2D1DeviceContext = unsafe { d2d_device.CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE)? };
            let dc5: ID2D1DeviceContext5 = dc.cast()?;
            *context_ref = Some(dc5);
        }
        Ok(context_ref.as_ref().unwrap().clone())
    })?;

    Ok((d2d_factory, d2d_device, d2d_context))
}

// RAII wrapper for D2D bitmap mapping - automatically unmaps when dropped
struct BitmapMapGuard<'a> {
    bitmap: &'a ID2D1Bitmap1,
    mapped: bool,
}

impl<'a> BitmapMapGuard<'a> {
    fn new(bitmap: &'a ID2D1Bitmap1) -> Result<(Self, D2D1_MAPPED_RECT)> {
        let mapped_rect = unsafe { bitmap.Map(D2D1_MAP_OPTIONS_READ)? };
        Ok((Self { bitmap, mapped: true }, mapped_rect))
    }
}

impl<'a> Drop for BitmapMapGuard<'a> {
    fn drop(&mut self) {
        if self.mapped {
            unsafe { let _ = self.bitmap.Unmap(); }
        }
    }
}

// RAII wrapper for D2D BeginDraw/EndDraw - automatically calls EndDraw when dropped
struct D2D1DrawGuard<'a> {
    context: &'a ID2D1DeviceContext5,
}

impl<'a> D2D1DrawGuard<'a> {
    fn new(context: &'a ID2D1DeviceContext5) -> Self {
        unsafe { context.BeginDraw() };
        Self { context }
    }
}

impl<'a> Drop for D2D1DrawGuard<'a> {
    fn drop(&mut self) {
        // Check the result of EndDraw. If the device is lost, poison the thread's resources so they will be recreated on the next run.
        let result = unsafe { self.context.EndDraw(None, None) };
        if let Err(e) = &result {
            if e.code() == D2DERR_RECREATE_TARGET {
                D2D_RESOURCES_POISONED.set(true);
            }
        }
    }
}

// RAII wrapper for VARIANT - automatically calls VariantClear when dropped
struct VariantGuard(VARIANT);

impl VariantGuard {
    fn new() -> Self {
        Self(VARIANT::default())
    }
}

impl Drop for VariantGuard {
    fn drop(&mut self) {
        // This is safe to call even on a default/zeroed VARIANT
        unsafe { let _ = VariantClear(&mut self.0); }
    }
}

impl std::ops::Deref for VariantGuard {
    type Target = VARIANT;
    fn deref(&self) -> &Self::Target {
        &self.0
    }
}

impl std::ops::DerefMut for VariantGuard {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.0
    }
}

/// Parses CSS text content and returns a list of class names and their concatenated style properties.
fn parse_css_rules(css_content: &str) -> Vec<(String, String)> {
    let mut style_list: Vec<(String, String)> = Vec::new();
    
    // Clean the input string: remove leading/trailing whitespace and control characters,
    // which should get rid of the junk data you're seeing from the buffer.
    let cleaned_content = remove_css_comments(css_content.trim());
    
    // Split by '}' to get individual rules
    let rules: Vec<&str> = cleaned_content.split('}').collect();
    
    for rule in rules {
        let rule = rule.trim();
        if rule.is_empty() {
            continue;
        }
        
        // Find the opening brace to separate selectors from properties
        if let Some(brace_pos) = rule.find('{') {
            let selectors_part = rule[..brace_pos].trim();
            let properties_part = rule[brace_pos + 1..].trim();
            
            // Split selectors by comma for grouped rules like ".cls-1, .cls-2"
            let selectors: Vec<&str> = selectors_part.split(',').collect();
            
            for selector in selectors {
                let selector = selector.trim();
                
                // We only care about class selectors (starting with '.')
                if let Some(class_name) = selector.strip_prefix('.') {
                    let class_name = class_name.trim().to_string();
                    let normalized_properties = normalize_css_properties(properties_part);

                    // Find if this class already exists in our list
                    if let Some((_key, existing_styles)) = style_list.iter_mut().find(|(key, _)| key == &class_name) {
                        // If yes, append the new properties
                        existing_styles.push_str(&normalized_properties);
                    } else {
                        // If no, add a new entry to the list
                        style_list.push((class_name, normalized_properties));
                    }
                }
            }
        }
    }
    
    style_list
}

/// Removes CSS comments from the input string.
fn remove_css_comments(css: &str) -> String {
    let mut result = String::new();
    let mut chars = css.chars().peekable();
    
    while let Some(ch) = chars.next() {
        if ch == '/' && chars.peek() == Some(&'*') {
            // Start of comment, consume until */
            chars.next(); // consume '*'
            
            while let Some(ch) = chars.next() {
                if ch == '*' && chars.peek() == Some(&'/') {
                    chars.next(); // consume '/'
                    break;
                }
            }
        } else {
            result.push(ch);
        }
    }
    
    result
}

/// Normalizes CSS properties to ensure they're properly formatted for inline styles.
fn normalize_css_properties(properties: &str) -> String {
    let mut result = String::new();
    
    // Split by semicolon and clean up each property
    let props: Vec<&str> = properties.split(';').collect();
    
    for prop in props {
        let prop = prop.trim();
        if !prop.is_empty() {
            result.push_str(prop);
            if !prop.ends_with(';') {
                result.push(';');
            }
        }
    }
    
    result
}

/// Parses raw SVG data to extract CSS from <style> tags found within a <defs> block.
/// This avoids the unreliable ID2D1SvgElement DOM traversal for styles.
/// Returns a single string containing all found CSS rules.
fn extract_css_from_svg_data(svg_data: &[u8]) -> String {
    let svg_string = String::from_utf8_lossy(svg_data);
    let mut css_content = String::new();

    // Find the <defs> block first. We only care about styles inside it.
    if let Some(defs_start) = svg_string.find("<defs") {
        if let Some(defs_end) = svg_string[defs_start..].find("</defs>") {
            let defs_block_content = &svg_string[defs_start..defs_start + defs_end];
            
            // Now, find all <style> tags within that block.
            let mut current_pos = 0;
            while let Some(style_start_tag) = defs_block_content[current_pos..].find("<style") {
                let style_start_abs = current_pos + style_start_tag;
                if let Some(style_content_start) = defs_block_content[style_start_abs..].find('>') {
                    let style_content_start_abs = style_start_abs + style_content_start + 1;
                    if let Some(style_content_end) = defs_block_content[style_content_start_abs..].find("</style>") {
                        let style_content_end_abs = style_content_start_abs + style_content_end;
                        
                        let content = &defs_block_content[style_content_start_abs..style_content_end_abs];
                        css_content.push_str(content);
                        css_content.push('\n'); // Add a newline for separation.
                        
                        current_pos = style_content_end_abs;
                    } else { break; } // No closing </style> tag found.
                } else { break; } // No opening '>' found for style tag.
            }
        }
    }
    
    css_content
}

/// Applies inline styles to SVG elements based on their class attributes using the MSXML parser.
/// It loads the SVG, finds elements by class, applies the provided styles, and returns the modified SVG data.
fn preprocess_svg_with_msxml(svg_data: &[u8], style_map: &[(String, String)]) -> Result<Vec<u8>> {
    // Skip it all if there are no styles to apply.
    if style_map.is_empty() {
        return Ok(svg_data.to_vec());
    }

    // MSXML is a COM (Component Object Model) library. Any thread that uses COM must first initialize it.
    // The `ComGuard` is an RAII wrapper that calls `CoInitializeEx` on creation and `CoUninitialize` on drop, ensuring cleanup.
    let _com_guard = ComGuard::new()?;

    // This creates an instance of the MSXML6 DOM Document object, which is our XML parser.
    // `CoCreateInstance` is the standard COM function for creating objects from a CLSID (Class ID).
    let dom: MsXml::IXMLDOMDocument2 = unsafe { Com::CoCreateInstance(&DOMDocument60, None, Com::CLSCTX_INPROC_SERVER)? };
    
    // --- Load SVG data into the DOM document ---

    // `IStream` is a standard COM interface for streamable data, behaving like an in-memory file.
    let stream: Com::IStream = unsafe { Shell::SHCreateMemStream(Some(svg_data)) }
        .ok_or_else(|| Error::new(E_FAIL, "Failed to create memory stream for MSXML"))?;
    
    // The `dom.load` method is particular and requires its input to be a `VARIANT` (a special COM struct that can hold many different types of data.)
    // We use our `VariantGuard` to ensure `VariantClear` is called, which will correctly release the COM object we're about to put in it.
    let mut stream_variant = VariantGuard::new();
    unsafe {
        // Get a mutable reference to the anonymous union inside the `VARIANT` struct.
        let v = &mut stream_variant.Anonymous.Anonymous;
        // Set variant type tag to `VT_UNKNOWN` -- it holds a generic COM interface pointer (`IUnknown`).
        v.vt = VT_UNKNOWN;
        // Transfer ownership of the `IStream` COM object to the `VARIANT`.
        // `stream.into()` converts the `IStream` smart pointer into its base `IUnknown` smart pointer.
        // `std::mem::ManuallyDrop::new` is CRITICAL: it prevents Rust from calling `Release` on the `stream` variable when it goes out of scope.
        // We have given ownership to the `VARIANT`, so the `VariantGuard` is now responsible for its cleanup. This prevents a double-release crash.
        v.Anonymous.punkVal = std::mem::ManuallyDrop::new(Some(stream.into()));
    }

    // The MSXML parser will read the SVG data directly from our in-memory stream.
    let success = unsafe { dom.load(&stream_variant)? };
    // For `dom.load`, success is specifically indicated by `VARIANT_TRUE` (-1), not just `S_OK` (0).
    if success != VARIANT_TRUE {
        return Err(Error::new(E_FAIL, "MSXML failed to load SVG data. It may be malformed."));
    }

    // --- Find elements with 'class' attribute and apply styles inline ---

    // `BSTR` is a length-prefixed, null-terminated wide string used by COM.
    let bstr_class = BSTR::from("class");
    let bstr_style = BSTR::from("style");

    // `selectNodes` uses an XPath query to find all elements in the document that have a "class" attribute.
    // This returns a collection of nodes that we can iterate over.
    let tagged_nodes: IXMLDOMNodeList = unsafe { dom.selectNodes(&BSTR::from("//*[@class]"))? };
    for i in 0..unsafe { tagged_nodes.length()? } {
        if let Ok(node) = unsafe { tagged_nodes.get_item(i) } {
            // A node could be a comment, text, etc. We only care about elements, so we try to cast it.
            // `cast` is a safe way to perform `QueryInterface` in `windows-rs`.
            if let Ok(element) = node.cast::<IXMLDOMElement>() {
                // Try to get the 'class' attribute from the current element.
                if let Ok(class_variant_raw) = unsafe { element.getAttribute(&bstr_class) } {
                    // `getAttribute` returns a new `VARIANT` which we now own. The guard ensures it's cleaned up.
                    let class_variant = VariantGuard(class_variant_raw);

                    let class_str = unsafe {
                        // We must check that the VARIANT actually contains a string (`BSTR`).
                        if (*class_variant.Anonymous.Anonymous).vt == VT_BSTR {
                            // This is the safest way to convert a `BSTR` inside a `VARIANT` to a Rust `String`.
                            // `(*...bstrVal)` gets the raw `BSTR` pointer from the `VARIANT`'s union. It is wrapped in `ManuallyDrop` by the bindings.
                            // We dereference it (`*`) to get a `&[u16]` slice of the raw string data without taking ownership.
                            // `String::from_utf16_lossy` then creates a new, Rust-owned `String` by *copying* the data from that slice.
                            // The original `BSTR` remains owned by the `VARIANT` and will be freed by the `VariantGuard`.
                            String::from_utf16_lossy(&(*class_variant.Anonymous.Anonymous).Anonymous.bstrVal)
                        } else {
                            String::new()
                        }
                    };

                    // If there's no class string, there's nothing to do for this element.
                    if class_str.is_empty() { continue; }

                    // This string will hold all the CSS rules for all classes on this element.
                    let mut combined_properties = String::new();
                    // An element can have multiple classes, e.g., `class="cls-1 cls-2"`. We split by whitespace to handle them all.
                    for class_name in class_str.split_whitespace() {
                        // Look up the current class name in our map of styles parsed from the `<style>` tag.
                        if let Some((_key, style_properties)) = style_map.iter().find(|(key, _)| key == class_name) {
                            // If found, append its CSS rules to our combined string.
                            combined_properties.push_str(style_properties);
                        }
                    }

                    // Only proceed if we actually found any styles to apply for the classes on this element.
                    if !combined_properties.is_empty() {
                        let mut existing_style = String::new();
                        // Check if the element *already* has an inline `style="..."` attribute.
                        if let Ok(style_variant_raw) = unsafe { element.getAttribute(&bstr_style) } {
                            let style_variant = VariantGuard(style_variant_raw);
                            // Also use the safe BSTR-to-String conversion for the style attribute.
                            if unsafe { (*style_variant.Anonymous.Anonymous).vt == VT_BSTR } {
                                existing_style = unsafe {
                                    String::from_utf16_lossy(&(*style_variant.Anonymous.Anonymous).Anonymous.bstrVal)
                                };
                                // To preserve existing styles, we need to append them. Ensure there's a semicolon separator.
                                if !existing_style.is_empty() && !existing_style.ends_with(';') {
                                    existing_style.push(';');
                                }
                            }
                        }

                        // Combine the new styles from the CSS classes with any pre-existing inline styles.
                        // We prepend our new styles so that existing inline styles can override them if needed, which is standard CSS behavior.
                        let final_style = format!("{}{}", combined_properties, existing_style);
                        // Create a new BSTR from our final combined Rust String.
                        let bstr = BSTR::from(final_style);
                        // We need to put this new BSTR into a VARIANT to pass it to `setAttribute`.
                        let mut variant_value = VariantGuard::new();
                        unsafe {
                            let v = &mut variant_value.Anonymous.Anonymous;
                            v.vt = VT_BSTR;
                            // Again, use `ManuallyDrop` to transfer ownership of the `bstr` to the `VARIANT`, preventing a double-free.
                            v.Anonymous.bstrVal = std::mem::ManuallyDrop::new(bstr);
                        }

                        // Finally, set the 'style' attribute on the element with our new, combined style string.
                        let _ = unsafe { element.setAttribute(&bstr_style, &variant_value) };
                    }
                }
            }
        }
    }

    // After the loop has modified the DOM in memory, serialize the entire document back into a BSTR string.
    let modified_xml_bstr = unsafe { dom.xml()? };
    // The `windows::core::BSTR` type is a smart pointer that will auto-free the string.
    let modified_xml_string = modified_xml_bstr.to_string();

    //DEBUG print the modified XML string log
    // log_message(modified_xml_string.as_str());

    // Convert the final string to a byte vector and return it.
    Ok(modified_xml_string.into_bytes())
}

pub fn render_svg_to_hbitmap(svg_data: &[u8], width: u32, height: u32) -> Result<Gdi::HBITMAP> {
    // Encapsulate main rendering logic in a helper closure.
    // This makes it easier to catch any error, check if it's D2DERR_RECREATE_TARGET, poison the resources if needed, and then return the original error.
    let result = (|| -> Result<Gdi::HBITMAP> {
        // Early validation - avoid work for invalid sizes
        if width == 0 || height == 0 || width > 4096 || height > 4096 {
            return Err(Error::new(E_INVALIDARG, "Invalid bitmap dimensions"));
        }

        // 1. Get resources (now includes cached device context)
        let (_d2d_factory, _d2d_device, d2d_context) = get_d2d_resources()?;
        
        // 2. Create the D2D RENDER TARGET bitmap (GPU-only)
        let bitmap_props_rt = D2D1_BITMAP_PROPERTIES1 {
            pixelFormat: D2D1_PIXEL_FORMAT { format: Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM, alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED },
            dpiX: 96.0,
            dpiY: 96.0,
            bitmapOptions: D2D1_BITMAP_OPTIONS_TARGET,
            ..Default::default()
        };
        let render_target_bitmap: ID2D1Bitmap1 = unsafe { d2d_context.CreateBitmap(D2D_SIZE_U { width, height }, None, 0, &bitmap_props_rt) }?;
        
        // 3. Set target and draw the SVG
        unsafe { d2d_context.SetTarget(&render_target_bitmap) };
        {
            let _draw_guard = D2D1DrawGuard::new(&d2d_context);
            
            // Clear to transparent black
            unsafe { d2d_context.Clear(Some(&D2D1_COLOR_F { r: 0.0, g: 0.0, b: 0.0, a: 0.0 })) };
            
            // Phase 1: Manually parse styles from the raw SVG data.
            let css_content = extract_css_from_svg_data(svg_data);
            let style_map = parse_css_rules(&css_content);
            // log_message(&format!("Style list contents: {:?}", style_map));

            // Preprocess the SVG to inline all CSS styles from the map.
            // This returns a new SVG data buffer with styles applied as inline `style` attributes.
            let processed_svg_data = preprocess_svg_with_msxml(svg_data, &style_map)?;

            // Load the PROCESSED svg data into a memory stream.
            let stream: Com::IStream = unsafe { Shell::SHCreateMemStream(Some(&processed_svg_data)) }.ok_or_else(|| Error::new(E_FAIL, "Failed to create memory stream"))?;
            
            // Create the SVG document from the stream of PROCESSED SVG data.
            let svg_doc: ID2D1SvgDocument = unsafe { d2d_context.CreateSvgDocument(
                &stream,
                D2D_SIZE_F { 
                    width: width as f32, 
                    height: height as f32
                }
            ) }?;
            
            // Phase 2 is no longer needed as styles are inlined in the data stream.
            
            // Get the root <svg> element from the document, so we can get or change the top level attributes such as width, height, viewbox, etc.
            if let Ok(root_element) = unsafe { svg_doc.GetRoot() } {
                // Apparently if there are no width and height attributes, DrawSvgDocument will automatically scale it to the viewbox, which we have set to the size of the bitmap/thumbnail
                // So we can just remove them from before drawing, and it will autoscale and fill the thumbnail.
                unsafe {
                    let _ = root_element.RemoveAttribute(w!("height"));
                    let _ = root_element.RemoveAttribute(w!("width"));
                }
            }
            
            unsafe { d2d_context.DrawSvgDocument(&svg_doc) };
        } // EndDraw called here by guard
        
        // Clear target before applying effects
        unsafe { d2d_context.SetTarget(None) };
        
        // Apply UnPremultiply effect
        let final_bitmap: ID2D1Bitmap1 = match unsafe { d2d_context.CreateEffect(&Direct2D::CLSID_D2D1UnPremultiply) } {
            Ok(unpremultiply_effect) => {
                // Create a second render target bitmap for the UnPremultiply effect output
                let output_bitmap: ID2D1Bitmap1 = unsafe { d2d_context.CreateBitmap(D2D_SIZE_U { width, height }, None, 0, &bitmap_props_rt) }?;
                
                // Switch to the output bitmap as the target and begin a new draw session
                unsafe { d2d_context.SetTarget(&output_bitmap) };
                {
                    let _effect_draw_guard = D2D1DrawGuard::new(&d2d_context);
                    
                    // SetInput doesn't return a Result, it's a void method
                    unsafe { unpremultiply_effect.SetInput(0, &render_target_bitmap, true) };
                    
                    match unpremultiply_effect.cast::<ID2D1Image>() {
                        Ok(effect_image) => {
                            // DrawImage doesn't return a Result either
                            unsafe { d2d_context.DrawImage(&effect_image, None, None, D2D1_INTERPOLATION_MODE_LINEAR, D2D1_COMPOSITE_MODE_SOURCE_COPY) };
                        }
                        Err(_) => {
                            // Effect cast failed, but we'll still return the output bitmap
                            // The draw guard will clean up properly
                        }
                    }
                } // EndDraw called here by guard
                
                // Clear target after effect drawing
                unsafe { d2d_context.SetTarget(None) };
                
                // Return the output bitmap from the UnPremultiply effect
                output_bitmap
            }
            Err(_) => {
                // Fall back to original bitmap if effect creation fails
                render_target_bitmap
            }
        };

        // 4. Create the CPU-readable STAGING bitmap
        let bitmap_props_staging = D2D1_BITMAP_PROPERTIES1 {
            pixelFormat: D2D1_PIXEL_FORMAT { format: Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM, alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED },
            dpiX: 96.0,
            dpiY: 96.0,
            bitmapOptions: D2D1_BITMAP_OPTIONS_CPU_READ | D2D1_BITMAP_OPTIONS_CANNOT_DRAW,
            ..Default::default()
        };
        let staging_bitmap: ID2D1Bitmap1 = unsafe { d2d_context.CreateBitmap(D2D_SIZE_U { width, height }, None, 0, &bitmap_props_staging) }?;
        
        // 5. Copy from render target to staging bitmap (GPU -> CPU accessible D2D memory)
        // This copies the pixel data but it's still in D2D's memory space
        unsafe { staging_bitmap.CopyFromBitmap(None, &final_bitmap, None) }?;
        
        // 6. Map the staging bitmap to get a pointer to the pixel data using RAII guard
        let (map_guard, mapped_rect) = BitmapMapGuard::new(&staging_bitmap)?;
        
        // 7. Create the final GDI HBITMAP
        // This creates a separate GDI bitmap with its own memory buffer
        let bmi = Gdi::BITMAPINFO { bmiHeader: Gdi::BITMAPINFOHEADER {
            biSize: std::mem::size_of::<Gdi::BITMAPINFOHEADER>() as u32, biWidth: width as i32, biHeight: -(height as i32),
            biPlanes: 1, biBitCount: 32, biCompression: Gdi::BI_RGB.0 as u32, ..Default::default()
        }, ..Default::default() };

        let mut dib_data: *mut std::ffi::c_void = std::ptr::null_mut();
        let hbitmap: Gdi::HBITMAP = unsafe { Gdi::CreateDIBSection(None, &bmi, Gdi::DIB_RGB_COLORS, &mut dib_data, None, 0) }?;
        
        // 8. Copy pixels from the mapped D2D buffer to the GDI HBITMAP buffer
        if !dib_data.is_null() {
            // Create safe slices from the raw pointers.
            let source_data: &[u8] = unsafe { std::slice::from_raw_parts(mapped_rect.bits, (mapped_rect.pitch * height) as usize) };
            let dest_data: &mut [u8] = unsafe { std::slice::from_raw_parts_mut(dib_data.cast::<u8>(), (width * height * 4) as usize) };
            
            // PRE-INITIALIZE the destination buffer to zero. This is the simplest way to prevent
            // garbage data in any padding bytes left over from a stride mismatch.
            dest_data.fill(0);

            // Now, copy the image data.
            if mapped_rect.pitch == (width * 4) {
                // Direct copy if stride matches.
                dest_data.copy_from_slice(&source_data[..dest_data.len()]);
            } else {
                // Copy row by row to handle stride differences.
                let dest_stride: usize = (width * 4) as usize;
                let source_stride: usize = mapped_rect.pitch as usize;
                let row_copy_len = std::cmp::min(dest_stride, source_stride);

                for y in 0..height as usize {
                    let src_start: usize = y * source_stride;
                    let dest_start: usize = y * dest_stride;
                    
                    let src_slice = &source_data[src_start .. src_start + row_copy_len];
                    let dest_slice = &mut dest_data[dest_start .. dest_start + row_copy_len];
                    dest_slice.copy_from_slice(src_slice);
                }
            }
        }

        // The map_guard will automatically unmap the bitmap when it goes out of scope
        drop(map_guard);
        
        Ok(hbitmap)
    })();

    // Check if the closure returned an error, and if that error was due to a lost device.
    // Set the poisoned flag if so, to force recreation of resources next time.
    if let Err(e) = &result {
        if e.code() == D2DERR_RECREATE_TARGET {
            D2D_RESOURCES_POISONED.set(true);
        }
    }

    result
}

// =================================================================
//                 COM Thumbnail Provider Object
// =================================================================

#[implement(Shell::PropertiesSystem::IInitializeWithStream, Shell::IThumbnailProvider)]
struct ThumbnailProvider {
    svg_data: Mutex<Option<Arc<[u8]>>>,
}

impl Default for ThumbnailProvider {
    fn default() -> Self {
        dll_add_ref();
        Self {
            svg_data: Mutex::new(None),
        }
    }
}

impl Drop for ThumbnailProvider {
    fn drop(&mut self) {
        dll_release();
    }
}

impl Shell::PropertiesSystem::IInitializeWithStream_Impl for ThumbnailProvider_Impl {
    #[allow(non_snake_case)]
    fn Initialize(&self, pstream: Ref<'_, Com::IStream>, _grfmode: u32) -> Result<()> {
        ffi_guard!(Result<()>, {
            // Guard against repeated initialization calls
            if self.svg_data.lock().map_err(|_| Error::new(E_FAIL, "Mutex was poisoned"))?.is_some() {
                return Err(Error::from(HRESULT::from_win32(ERROR_ALREADY_INITIALIZED.0)));
            }

            match &*pstream {
                Some(stream) => {
                    // 101 MiB max file size.
                    const MAX_SIZE: u64 = 101 * 1024 * 1024;
                    pub const ERROR_FILE_TOO_LARGE: WIN32_ERROR = WIN32_ERROR(223u32);

                    // Fast Fail Check: Ask the stream for its size for a quick rejection.
                    // If the size check fails continue to read the stream in chunks, there is another safety net below.
                    let mut statstg = Default::default();
                    if unsafe { stream.Stat(&mut statstg, Com::STATFLAG_NONAME) }.is_ok() {
                        let stream_size = statstg.cbSize;
                        if stream_size > 0 && stream_size > MAX_SIZE {
                            return Err(Error::from(HRESULT::from_win32(ERROR_FILE_TOO_LARGE.0)));
                        }
                    }

                    // Do not trust the reported size for allocation.
                    // Start with a default-sized Vec and let it grow.
                    let seq_stream: Com::ISequentialStream = stream.cast()?;
                    let mut buffer: Vec<u8> = Vec::new();
                    let mut chunk: Vec<u8> = vec![0u8; 65536];
                    
                    loop {
                        let mut bytes_read: u32 = 0;
                        let hr: HRESULT = unsafe {
                            seq_stream.Read(
                                chunk.as_mut_ptr() as *mut core::ffi::c_void,
                                chunk.len() as u32,
                                Some(&mut bytes_read)
                            )
                        };
                        
                        if hr.is_err() || bytes_read == 0 {
                            break;
                        }
                        
                        // Extra file size safety net protects memory usage in case statstg failed or returned a wrong size.
                        if buffer.len() + (bytes_read as usize) > (MAX_SIZE as usize) {
                            return Err(Error::from(HRESULT::from_win32(ERROR_FILE_TOO_LARGE.0)));
                        }
                        
                        buffer.extend_from_slice(&chunk[..bytes_read as usize]);
                    }
                    
                    // Convert to Arc<[u8]> to save memory overhead
                    *self.svg_data.lock().map_err(|_| Error::new(E_FAIL, "Mutex was poisoned"))? = Some(Arc::from(buffer.into_boxed_slice()));
                    
                    //log_message("Initialize: Succeeded.");
                    Ok(())
                }
                None => {
                    // This case handles if Windows passes a null stream.
                    //log_message("Initialize: Error - Stream was null.");
                    Err(E_INVALIDARG.into())
                }
            }
        })
    }
}

impl Shell::IThumbnailProvider_Impl for ThumbnailProvider_Impl {
    #[allow(non_snake_case)]
    fn GetThumbnail(&self, cx: u32, phbmp: *mut Gdi::HBITMAP, pdwalpha: *mut Shell::WTS_ALPHATYPE) -> Result<()> {
        ffi_guard!(Result<()>, {
            //log_message("GetThumbnail: Entered.");

            // Initialize output parameters to safe defaults (COM contract requirement)
            // pdwalpha is set to UNKNOWN for all failure cases, only changed to ARGB on success
            unsafe {
                *phbmp = Gdi::HBITMAP(std::ptr::null_mut());
                *pdwalpha = Shell::WTSAT_UNKNOWN;
            }

            // Clone the Arc (cheap pointer copy) and release the mutex before rendering to prevent deadlocks
            let svg_data = {
                let data_guard = self.svg_data.lock().map_err(|_| Error::new(E_FAIL, "Mutex was poisoned"))?;
                
                match data_guard.as_ref() {
                    Some(data) => {
                        //log_message(&format!("GetThumbnail: SVG data is {} bytes.", data.len()));
                        Arc::clone(data) // Clone the Arc (cheap pointer copy)
                    }
                    None => {
                        //log_message("GetThumbnail: Error - SVG data was not initialized.");
                        return Err(Error::new(E_UNEXPECTED, "SVG data not initialized"));
                    }
                }
            }; // Mutex lock is released here

            match render_svg_to_hbitmap(&svg_data[..], cx, cx) {
                Ok(hbitmap) => {
                    //log_message("GetThumbnail: render_svg_to_hbitmap succeeded.");
                    unsafe {
                        *phbmp = hbitmap;
                        *pdwalpha = Shell::WTSAT_ARGB;
                    }
                    //log_message("GetThumbnail: Succeeded.");
                    Ok(())
                }
                Err(e) => {
                    //log_message(&format!("GetThumbnail: render_svg_to_hbitmap failed with error: {:?}", e));
                    Err(e)
                }
            }
        })
    }
}

// // -------------- Logger ----------------
// fn log_message(message: &str) {
//     if let Ok(mut file) = std::fs::OpenOptions::new()
//         .create(true)
//         .append(true)
//         .open("C:\\temp\\svg_thumb_log.txt") // Make sure C:\temp exists!
//     {
//         let time = std::time::SystemTime::now()
//             .duration_since(std::time::UNIX_EPOCH)
//             .unwrap_or_default()
//             .as_secs();
//         let _ = writeln!(file, "[{}] {}", time, message);
//     }
// }

// fn log_message(message: &str) {
//     println!("{}", message);
// }

// =================================================================
//                      COM Class Factory
// =================================================================

#[implement(Com::IClassFactory)]
struct ClassFactory;

impl Default for ClassFactory {
    fn default() -> Self {
        dll_add_ref();
        Self {}
    }
}

impl Drop for ClassFactory {
    fn drop(&mut self) {
        dll_release();
    }
}

impl Com::IClassFactory_Impl for ClassFactory_Impl {
    #[allow(non_snake_case)]
    fn CreateInstance(&self, punkouter: Ref<'_, IUnknown>, riid: *const GUID, ppvobject: *mut *mut std::ffi::c_void) -> Result<()> {
        ffi_guard!(Result<()>, {
            //log_message(&format!("ClassFactory::CreateInstance: Entered. Requesting interface: {:?}", unsafe { *riid }));

            // Safety checks for null pointers
            if riid.is_null() || ppvobject.is_null() {
                return Err(Error::new(E_POINTER, "Null pointer passed to CreateInstance"));
            }

            // We do not support aggregation.
            if !punkouter.is_null() {
                //log_message("ClassFactory::CreateInstance: Error - Aggregation not supported.");
                return Err(Error::new(CLASS_E_NOAGGREGATION, "Aggregation not supported"));
            }
            
            // Create an instance of our ThumbnailProvider
            let thumbnail_provider: IUnknown = ThumbnailProvider::default().into();
            
            // Query for the interface requested by the caller and return it.
            let hr: HRESULT = unsafe { thumbnail_provider.query(&*riid, ppvobject) };

            //log_message(&format!("ClassFactory::CreateInstance: Exiting with HRESULT: {:?}", hr));
            
            if hr.is_ok() {
                Ok(())
            } else {
                Err(Error::new(hr, "Failed to query interface"))
            }
        })
    }

    #[allow(non_snake_case)]
    fn LockServer(&self, flock: BOOL) -> Result<()> {
        ffi_guard!(Result<()>, {
            if flock.as_bool() {
                dll_add_ref();
            } else {
                dll_release();
            }
            Ok(())
        })
    }
}

// =================================================================
//                      DLL Global State & Exports
// =================================================================

// A global reference counter for the DLL itself.
static DLL_REFERENCES: AtomicU32 = AtomicU32::new(0);
// A global handle to the DLL module instance - using Option for safer null checking
static MODULE_HANDLE: AtomicPtr<std::ffi::c_void> = AtomicPtr::new(std::ptr::null_mut());

fn dll_add_ref() {
    DLL_REFERENCES.fetch_add(1, Ordering::Relaxed);
}
fn dll_release() {
    DLL_REFERENCES.fetch_sub(1, Ordering::Release);
}

// This is our thumbnail provider's unique Class ID (CLSID).
// Use a new GUID for your own projects!
const CLSID_SVG_THUMBNAIL_PROVIDER: GUID = GUID::from_u128(0x95724385_3234_4ea4_8086_3499F447884D);

#[no_mangle]
#[allow(non_snake_case)]
extern "system" fn DllMain(hinst_dll: HMODULE, fdw_reason: u32, _lpv_reserved: *const std::ffi::c_void) -> BOOL {
    ffi_guard!(BOOL, {
        if fdw_reason == System::SystemServices::DLL_PROCESS_ATTACH {
            //log_message("DllMain: DLL_PROCESS_ATTACH received. DLL is loaded.");
            MODULE_HANDLE.store(hinst_dll.0 as *mut _, Ordering::Release);
        }
        true
    })
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllGetClassObject(rclsid: *const GUID, riid: *const GUID, ppv: *mut *mut std::ffi::c_void) -> HRESULT {
    ffi_guard!(HRESULT, {
        // Safety checks for null pointers
        if rclsid.is_null() || riid.is_null() || ppv.is_null() {
            return E_POINTER;
        }

        // Check if the caller is asking for our specific class.
        if unsafe { *rclsid } != CLSID_SVG_THUMBNAIL_PROVIDER {
            //log_message(&format!("DllGetClassObject: Error - CLSID mismatch. Requested: {:?}, Expected: {:?}", unsafe { *rclsid }, CLSID_SVG_THUMBNAIL_PROVIDER));
            return CLASS_E_CLASSNOTAVAILABLE;
        }
        
        // Create our class factory.
        let factory: Com::IClassFactory = ClassFactory::default().into();
        
        // Query for the interface the caller wants (usually IClassFactory) and return it.
        let hr: HRESULT = unsafe { factory.query(riid, ppv) };
        
        // The factory variable will automatically drop here, releasing our local reference.
        // The caller retains their reference from the query() call.

        //log_message(&format!("DllGetClassObject: Exiting with HRESULT: {:?}", hr));
        
        hr
    })
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllCanUnloadNow() -> HRESULT {
    ffi_guard!(HRESULT, {
        if DLL_REFERENCES.load(Ordering::Acquire) == 0 {
            S_OK
        } else {
            S_FALSE
        }
    })
}


// =================================================================
//                      DLL Registration
// =================================================================

/// Helper to convert a Rust string slice to a null-terminated UTF-16 wide string.
fn to_pcwstr(s: &str) -> Vec<u16> {
    OsStr::new(s).encode_wide().chain(std::iter::once(0)).collect()
}

fn create_registry_keys() -> Result<()> {
    let clsid_string = format!("{{{CLSID_SVG_THUMBNAIL_PROVIDER:?}}}");
    let dll_path: String = get_dll_path()?;

    // Prepare string values outside unsafe block
    let clsid_wide = to_pcwstr(&clsid_string);
    let value = to_pcwstr("SVG Thumbnail Provider (Rust)");
    let path_value = to_pcwstr(&dll_path);
    let model_value = to_pcwstr("Apartment");
    let clsid_value = to_pcwstr(&clsid_string);

    unsafe {
        // Create CLSID\{our-clsid} - using RAII wrapper for automatic cleanup
        let clsid_root_key = {
            let mut key = HKEY::default();
            let mut disposition = REG_CREATE_KEY_DISPOSITION(0);
            RegCreateKeyExW(
                HKEY_CLASSES_ROOT,
                w!("CLSID"),
                None,
                None,
                REG_OPTION_NON_VOLATILE,
                WRITE_FLAGS,
                None,
                &mut key,
                Some(&mut disposition as *mut _)
            ).ok()?;
            RegistryKeyGuard(key)
        };
        
        let clsid_key = clsid_root_key.create_subkey(&PCWSTR(clsid_wide.as_ptr()))?;
        RegSetValueExW(clsid_key.get(), PCWSTR::null(), Some(0), REG_SZ, Some(std::slice::from_raw_parts(value.as_ptr() as *const u8, value.len() * 2))).ok()?;

        // Create CLSID\{our-clsid}\InprocServer32
        let inproc_key = clsid_key.create_subkey(&w!("InprocServer32"))?;
        RegSetValueExW(inproc_key.get(), PCWSTR::null(), Some(0), REG_SZ, Some(std::slice::from_raw_parts(path_value.as_ptr() as *const u8, path_value.len() * 2))).ok()?;
        RegSetValueExW(inproc_key.get(), w!("ThreadingModel"), Some(0), REG_SZ, Some(std::slice::from_raw_parts(model_value.as_ptr() as *const u8, model_value.len() * 2))).ok()?;

        // Associate with .svg files by creating the key path explicitly in the correct registry view
        let svg_root_key = RegistryKeyGuard(HKEY_CLASSES_ROOT).create_subkey(&w!(".svg"))?;
        let svg_shellex_key = svg_root_key.create_subkey(&w!("shellex"))?;
        let svg_handler_key = svg_shellex_key.create_subkey(&w!("{E357FCCD-A995-4576-B01F-234630154E96}"))?;
        RegSetValueExW(svg_handler_key.get(), PCWSTR::null(), Some(0), REG_SZ, Some(std::slice::from_raw_parts(clsid_value.as_ptr() as *const u8, clsid_value.len() * 2))).ok()?;

        // Associate with .svgz files
        let svgz_root_key = RegistryKeyGuard(HKEY_CLASSES_ROOT).create_subkey(&w!(".svgz"))?;
        let svgz_shellex_key = svgz_root_key.create_subkey(&w!("shellex"))?;
        let svgz_handler_key = svgz_shellex_key.create_subkey(&w!("{E357FCCD-A995-4576-B01F-234630154E96}"))?;
        RegSetValueExW(svgz_handler_key.get(), PCWSTR::null(), Some(0), REG_SZ, Some(std::slice::from_raw_parts(clsid_value.as_ptr() as *const u8, clsid_value.len() * 2))).ok()?;

        Shell::SHChangeNotify(Shell::SHCNE_ASSOCCHANGED, Shell::SHCNF_IDLIST, None, None);
    }

    Ok(())
}

fn get_dll_path() -> Result<String> {
    let handle_ptr: *mut std::ffi::c_void = MODULE_HANDLE.load(Ordering::Acquire);

    // Check for null pointer to avoid potential crashes
    if handle_ptr.is_null() {
        return Err(Error::new(E_FAIL, "MODULE_HANDLE is null; DLL not loaded?"));
    }

    let handle: HMODULE = HMODULE(handle_ptr);
    let mut path = vec![0u16; MAX_PATH as usize];
    let len: u32 = unsafe { System::LibraryLoader::GetModuleFileNameW(Some(handle), &mut path) };

    // If the returned length is zero, it's an error
    if len == 0 {
        return Err(Error::new(E_FAIL, "GetModuleFileNameW failed (returned 0)"));
    }

    // If the returned length is equal to the buffer size, truncation may have occurred
    if (len as usize) >= path.len() {
        return Err(Error::new(E_FAIL, "DLL path is too long (truncated); registration aborted"));
    }

    // Additional safety check - ensure we don't go beyond the buffer
    let len = std::cmp::min(len as usize, path.len());
    Ok(String::from_utf16_lossy(&path[..len]))
}

// RAII wrapper for registry keys - automatically closes when dropped
struct RegistryKeyGuard(HKEY);

impl Drop for RegistryKeyGuard {
    fn drop(&mut self) {
        if !self.0.is_invalid() {
            unsafe { let _ = RegCloseKey(self.0); }
        }
    }
}

impl RegistryKeyGuard {
    fn create_subkey(&self, name: &PCWSTR) -> Result<RegistryKeyGuard> {
        let mut key = HKEY::default();
        let mut disposition = REG_CREATE_KEY_DISPOSITION(0);
        unsafe {
            RegCreateKeyExW(
                self.0,
                *name,
                None,
                None,
                REG_OPTION_NON_VOLATILE,
                WRITE_FLAGS,
                None,
                &mut key,
                Some(&mut disposition as *mut _)
            ).ok()?;
            if key.is_invalid() {
                return Err(Error::new(E_FAIL, "RegCreateKeyExW returned null handle"));
            }
        }
        Ok(RegistryKeyGuard(key))
    }
    
    fn get(&self) -> HKEY {
        self.0
    }
}

fn delete_registry_keys() -> Result<()> {
    let clsid_string = format!("{{{CLSID_SVG_THUMBNAIL_PROVIDER:?}}}");

    // Prepare string values outside unsafe block
    let clsid_path = to_pcwstr(&format!("CLSID\\{}", clsid_string));

    unsafe {
        RegDeleteTreeW(HKEY_CLASSES_ROOT, PCWSTR(clsid_path.as_ptr())).ok()?;
        RegDeleteTreeW(HKEY_CLASSES_ROOT, w!(".svg\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}")).ok()?;
        RegDeleteTreeW(HKEY_CLASSES_ROOT, w!(".svgz\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}")).ok()?;

        Shell::SHChangeNotify(Shell::SHCNE_ASSOCCHANGED, Shell::SHCNF_IDLIST, None, None)
    }

    Ok(())
}


#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllRegisterServer() -> HRESULT {
    ffi_guard!(HRESULT, {
        match create_registry_keys() {
            Ok(_) => S_OK,
            Err(_) => E_FAIL,
        }
    })
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllUnregisterServer() -> HRESULT {
    ffi_guard!(HRESULT, {
        match delete_registry_keys() {
            Ok(_) => S_OK,
            Err(_) => E_FAIL,
        }
    })
}
