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
            Registry::*
        },
        UI::Shell
    },
};

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
}
/// Initializes and retrieves the thread-local Direct2D and WIC resources.
/// This function ensures that the heavyweight factory and device objects are created only once per thread.
fn get_d2d_resources() -> Result<(ID2D1Factory1, ID2D1Device, ID2D1DeviceContext5)> {
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
            let dc: ID2D1DeviceContext = unsafe { d2d_device.CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_ENABLE_MULTITHREADED_OPTIMIZATIONS)? };
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
        // Ignore errors in drop - we can't handle them anyway
        unsafe { let _ = self.context.EndDraw(None, None); }
    }
}

/// Renders SVG data to a GDI HBITMAP with an alpha channel using a robust staging bitmap technique.
pub fn render_svg_to_hbitmap(svg_data: &[u8], width: u32, height: u32) -> Result<Gdi::HBITMAP> {
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

        // Load svg data into a memory stream as the input for the SVG document
        let stream: Com::IStream = unsafe { Shell::SHCreateMemStream(Some(svg_data)) }.ok_or_else(|| Error::new(E_FAIL, "Failed to create memory stream"))?;

        // Create the SVG document from the stream of SVG data
        let svg_doc: ID2D1SvgDocument = unsafe { d2d_context.CreateSvgDocument(
            &stream,
            D2D_SIZE_F { 
                width: width as f32, 
                height: height as f32
            }
        ) }?;

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
    // Even though CopyFromBitmap moved data to CPU-accessible memory, it's still in D2D's memory space. We need to copy it to the GDI bitmap's memory.
    if !dib_data.is_null() {
        // Create safe slices from the raw pointers
        let source_data: &[u8] = unsafe { std::slice::from_raw_parts(mapped_rect.bits, (mapped_rect.pitch * height) as usize) };
        let dest_data: &mut [u8] = unsafe { std::slice::from_raw_parts_mut(dib_data.cast::<u8>(), (width * height * 4) as usize) };
        
        // Since the UnPremultiply effect already handled the alpha conversion, we can do a simple memory copy if the stride matches
        if mapped_rect.pitch == (width * 4) {
            // Direct copy - no stride mismatch
            dest_data.copy_from_slice(&source_data[..dest_data.len()]);
        } else {
            // Copy row by row to handle stride differences
            let dest_stride: usize = (width * 4) as usize;
            let source_stride: usize = mapped_rect.pitch as usize;
            
            for y in 0..height as usize {
                let src_start: usize = y * source_stride;
                let dest_start: usize = y * dest_stride;
                dest_data[dest_start..dest_start + dest_stride]
                    .copy_from_slice(&source_data[src_start..src_start + dest_stride]);
            }
        }
    }

    // The map_guard will automatically unmap the bitmap when it goes out of scope
    drop(map_guard);

    Ok(hbitmap)
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
            //log_message("Initialize: Entered.");

            // Guard against repeated initialization calls
            if self.svg_data.lock().map_err(|_| Error::new(E_FAIL, "Mutex was poisoned"))?.is_some() {
                return Err(Error::from(HRESULT::from_win32(ERROR_ALREADY_INITIALIZED.0)));
            }

            match &*pstream {
                Some(stream) => {
                    //log_message("Initialize: Stream is valid. Proceeding to read.");

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

// -------------- Logger ----------------
// fn log_message(message: &str) {
//     if let Ok(mut file) = OpenOptions::new()
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
    let dll_path: String = get_dll_path();

    // Prepare string values outside unsafe block
    let clsid_wide = to_pcwstr(&clsid_string);
    let value = to_pcwstr("SVG Thumbnail Provider (Rust)");
    let path_value = to_pcwstr(&dll_path);
    let model_value = to_pcwstr("Both");
    let clsid_value = to_pcwstr(&clsid_string);

    unsafe {
        // Create CLSID\{our-clsid} - using RAII wrapper for automatic cleanup
        let clsid_root_key = {
            let mut key = HKEY::default();
            RegCreateKeyW(HKEY_CLASSES_ROOT, w!("CLSID"), &mut key).ok()?;
            RegistryKeyGuard(key)
        };
        
        let clsid_key = clsid_root_key.create_subkey(&PCWSTR(clsid_wide.as_ptr()))?;
        RegSetValueExW(clsid_key.get(), PCWSTR::null(), Some(0), REG_SZ, Some(std::slice::from_raw_parts(value.as_ptr() as *const u8, value.len() * 2))).ok()?;

        // Create CLSID\{our-clsid}\InprocServer32
        let inproc_key = clsid_key.create_subkey(&w!("InprocServer32"))?;
        RegSetValueExW(inproc_key.get(), PCWSTR::null(), Some(0), REG_SZ, Some(std::slice::from_raw_parts(path_value.as_ptr() as *const u8, path_value.len() * 2))).ok()?;
        RegSetValueExW(inproc_key.get(), w!("ThreadingModel"), Some(0), REG_SZ, Some(std::slice::from_raw_parts(model_value.as_ptr() as *const u8, model_value.len() * 2))).ok()?;

        // Associate with .svg files
        let svg_key = {
            let mut key = HKEY::default();
            RegCreateKeyW(HKEY_CLASSES_ROOT, w!(".svg\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}"), &mut key).ok()?;
            RegistryKeyGuard(key)
        };
        RegSetValueExW(svg_key.get(), PCWSTR::null(), Some(0), REG_SZ, Some(std::slice::from_raw_parts(clsid_value.as_ptr() as *const u8, clsid_value.len() * 2))).ok()?;

        // Associate with .svgz files
        let svgz_key = {
            let mut key = HKEY::default();
            RegCreateKeyW(HKEY_CLASSES_ROOT, w!(".svgz\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}"), &mut key).ok()?;
            RegistryKeyGuard(key)
        };
        RegSetValueExW(svgz_key.get(), PCWSTR::null(), Some(0), REG_SZ, Some(std::slice::from_raw_parts(clsid_value.as_ptr() as *const u8, clsid_value.len() * 2))).ok()?;

        Shell::SHChangeNotify(Shell::SHCNE_ASSOCCHANGED, Shell::SHCNF_IDLIST, None, None);
    }

    Ok(())
}

fn get_dll_path() -> String {
    let handle_ptr: *mut std::ffi::c_void = MODULE_HANDLE.load(Ordering::Acquire);
    
    // Check for null pointer to avoid potential crashes
    if handle_ptr.is_null() {
        return String::new(); // Return empty string if handle is invalid
    }
    
    let handle: HMODULE = HMODULE(handle_ptr);
    let mut path = vec![0u16; MAX_PATH as usize];
    let len: u32 = unsafe { System::LibraryLoader::GetModuleFileNameW(Some(handle), &mut path) };
    
    // Additional safety check - ensure we don't go beyond the buffer
    let len = std::cmp::min(len as usize, path.len());
    String::from_utf16_lossy(&path[..len])
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
        unsafe { RegCreateKeyW(self.0, *name, &mut key).ok()? };
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
