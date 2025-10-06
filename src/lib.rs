use std::{
    ffi::OsStr,
    // fs::OpenOptions,
    io::Write,
    os::windows::prelude::OsStrExt,
    panic::{catch_unwind, AssertUnwindSafe},
    sync::{
        atomic::{
            AtomicPtr,
            AtomicU32,
            Ordering
        },
        Arc,
        Mutex,
        OnceLock
    },
    path::PathBuf,
};

use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::{
            self,
            Gdi,
        },
        System::{
            self,
            Com,
            Registry::{
                *,
                RegCreateKeyExW,
                RegSetValueExW,
            },
            SystemInformation::GetLocalTime
        },
        UI::Shell::{
            self,
            SHGetKnownFolderPath,
            FOLDERID_Desktop
        },
        Globalization::{GetTimeFormatEx, TIME_FORMAT_FLAGS},
    },
};

// This is the ONLY definition you need. It works for both 32-bit and 64-bit.
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
                //RESOURCES.with(|resources| {
                //    resources.borrow_mut().take();
                //});
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
                //RESOURCES.with(|resources| {
                //    resources.borrow_mut().take();
                //});
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
                //RESOURCES.with(|resources| {
                //    resources.borrow_mut().take();
                //});
                //log_message("A PANIC occurred in FFI function.");
                false.into()
            }
        }
    }};
}

// RAII wrapper for HBITMAP - automatically calls DeleteObject when dropped
struct HBitmapGuard(Gdi::HBITMAP);

impl HBitmapGuard {
    // Create a new guard. Takes ownership of the handle.
    fn new(handle: Gdi::HBITMAP) -> Self {
        Self(handle)
    }

    // Release ownership of the handle (e.g., when transferring to the Shell).
    fn release(mut self) -> Gdi::HBITMAP {
        let handle = self.0;
        // Set the internal handle to null so Drop doesn't delete it.
        self.0 = Gdi::HBITMAP(std::ptr::null_mut());
        handle
    }
}

impl Drop for HBitmapGuard {
    fn drop(&mut self) {
        // Only delete if the handle is not null/invalid (i.e., it hasn't been released).
        if !self.0.is_invalid() && !self.0.0.is_null() {
            let success = unsafe { Graphics::Gdi::DeleteObject(Gdi::HGDIOBJ(self.0.0)) };
            // In a Drop implementation, we can't return an error, so just ignore failure
            // In production, you might want to log this for debugging
            let _ = success;
        }
    }
}

// RAII Guard to automatically free memory from SHGetKnownFolderPath
struct CoTaskMemFreeGuard(PWSTR);

impl Drop for CoTaskMemFreeGuard {
    fn drop(&mut self) {
        if !self.0.is_null() {
            // This is safe because the pointer was allocated by the COM task allocator
            unsafe { Com::CoTaskMemFree(Some(self.0 .0 as *const std::ffi::c_void)) };
        }
    }
}

pub fn render_svg_to_hbitmap(svg_data: &[u8], requested_width: u32, requested_height: u32) -> Result<Gdi::HBITMAP> {
    log_message(&format!("render_svg_to_hbitmap: Starting render for {}x{} size, {} bytes of data", requested_width, requested_height, svg_data.len()));

    // 7. Create the final GDI HBITMAP
    // This creates a separate GDI bitmap with its own memory buffer
    let bmi = Gdi::BITMAPINFO { bmiHeader: Gdi::BITMAPINFOHEADER {
        biSize: std::mem::size_of::<Gdi::BITMAPINFOHEADER>() as u32, biWidth: requested_width as i32, biHeight: -(requested_height as i32),
        biPlanes: 1, biBitCount: 32, biCompression: Gdi::BI_RGB.0 as u32, ..Default::default()
    }, ..Default::default() };

    let mut dib_data: *mut std::ffi::c_void = std::ptr::null_mut();
    let hbitmap_handle: Gdi::HBITMAP = unsafe {
        Gdi::CreateDIBSection(None, &bmi, Gdi::DIB_RGB_COLORS, &mut dib_data, None, 0)
    }?;
    let hbitmap_guard = HBitmapGuard::new(hbitmap_handle);

    // 8. Copy pixels from the mapped D2D buffer to the GDI HBITMAP buffer
    if !dib_data.is_null() {
        // Safety: The bitmap bit values are aligned on doubleword boundaries
        let pixels = dib_data as *mut u32;

        // TODO: Draw pixels

    }

    log_message("render_svg_to_hbitmap: Successfully completed rendering");
    Ok(hbitmap_guard.release())
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
        log_message("ThumbnailProvider: Created new instance");
        Self {
            svg_data: Mutex::new(None),
        }
    }
}

impl Drop for ThumbnailProvider {
    fn drop(&mut self) {
        log_message("ThumbnailProvider: Dropping instance");
        dll_release();
    }
}

impl Shell::PropertiesSystem::IInitializeWithStream_Impl for ThumbnailProvider_Impl {
    #[allow(non_snake_case)]
    fn Initialize(&self, pstream: Ref<'_, Com::IStream>, _grfmode: u32) -> Result<()> {
        ffi_guard!(Result<()>, {
            // log_message("Initialize: Starting SVG data loading");

            // Guard against repeated initialization calls
            if self.svg_data.lock().map_err(|_| Error::new(E_FAIL, "Mutex was poisoned"))?.is_some() {
                log_message("Initialize: Error - Already initialized");
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
                        // log_message(&format!("Initialize: Stream reports size: {} bytes", stream_size));
                        if stream_size > 0 && stream_size > MAX_SIZE {
                            log_message(&format!("Initialize: Error - File too large: {} bytes (max: {} bytes)", stream_size, MAX_SIZE));
                            return Err(Error::from(HRESULT::from_win32(ERROR_FILE_TOO_LARGE.0)));
                        }
                    } else {
                        log_message("Initialize: Warning - Could not get stream size, will read with safety checks");
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
                            if hr.is_err() {
                                log_message(&format!("Initialize: Stream read error: {:?}", hr));
                            }
                            break;
                        }

                        // Extra file size safety net protects memory usage in case statstg failed or returned a wrong size.
                        if buffer.len() + (bytes_read as usize) > (MAX_SIZE as usize) {
                            log_message(&format!("Initialize: Error - File too large during read: {} bytes (max: {} bytes)", buffer.len() + (bytes_read as usize), MAX_SIZE));
                            return Err(Error::from(HRESULT::from_win32(ERROR_FILE_TOO_LARGE.0)));
                        }

                        buffer.extend_from_slice(&chunk[..bytes_read as usize]);
                    }

                    // log_message(&format!("Initialize: Successfully loaded {} bytes of SVG data", buffer.len()));

                    // Convert to Arc<[u8]> to save memory overhead
                    *self.svg_data.lock().map_err(|_| Error::new(E_FAIL, "Mutex was poisoned"))? = Some(Arc::from(buffer.into_boxed_slice()));

                    // log_message("Initialize: Succeeded.");
                    Ok(())
                }
                None => {
                    // This case handles if Windows passes a null stream.
                    log_message("Initialize: Error - Stream was null.");
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
            // log_message(&format!("GetThumbnail: Entered with size: {}x{}", cx, cx));

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
                        // log_message(&format!("GetThumbnail: SVG data is {} bytes.", data.len()));
                        Arc::clone(data) // Clone the Arc (cheap pointer copy)
                    }
                    None => {
                        log_message("GetThumbnail: Error - SVG data was not initialized.");
                        return Err(Error::new(E_UNEXPECTED, "SVG data not initialized"));
                    }
                }
            }; // Mutex lock is released here

            match render_svg_to_hbitmap(&svg_data[..], cx, cx) {
                Ok(hbitmap) => {
                    // log_message("GetThumbnail: render_svg_to_hbitmap succeeded.");
                    unsafe {
                        *phbmp = hbitmap;
                        *pdwalpha = Shell::WTSAT_ARGB;
                    }
                    // log_message("GetThumbnail: Succeeded.");
                    Ok(())
                }
                Err(e) => {
                    log_message(&format!("GetThumbnail: render_svg_to_hbitmap failed with error: {:?}", e));

                    // Instead of returning an error, create a fallback thumbnail
                    match create_fallback_thumbnail(cx) {
                        Ok(fallback_hbitmap) => {
                            log_message("GetThumbnail: Created fallback thumbnail for invalid SVG.");
                            unsafe {
                                *phbmp = fallback_hbitmap;
                                *pdwalpha = Shell::WTSAT_ARGB;
                            }
                            Ok(())
                        }
                        Err(fallback_err) => {
                            log_message(&format!("GetThumbnail: Failed to create fallback thumbnail: {:?}", fallback_err));
                            Err(e) // Only return error if we can't even create a fallback
                        }
                    }
                }
            }
        })
    }
}

/// Creates a simple fallback thumbnail for invalid SVG files
fn create_fallback_thumbnail(size: u32) -> Result<Gdi::HBITMAP> {
    // log_message(&format!("create_fallback_thumbnail: Creating fallback thumbnail of size {}x{}", size, size));

    // Use a hardcoded "broken file" SVG with red X pattern
    const FALLBACK_SVG: &[u8] = b"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 256 256\"><g><line stroke-width=\"2\" stroke=\"#ff0000\" y2=\"256\" x2=\"0\" y1=\"0\" x1=\"256\" fill=\"none\"/><line stroke-width=\"2\" y2=\"256\" x2=\"256\" y1=\"0\" x1=\"0\" stroke=\"#ff0000\" fill=\"none\"/></g></svg>";

    // Try to render the fallback SVG using our normal rendering pipeline
    match render_svg_to_hbitmap(FALLBACK_SVG, size, size) {
        Ok(hbitmap) => {
            log_message("create_fallback_thumbnail: Successfully created SVG-based fallback");
            Ok(hbitmap)
        },
        Err(e) => {
            log_message(&format!("create_fallback_thumbnail: SVG fallback failed: {:?}, creating bitmap fallback", e));
            // If even the fallback SVG fails to render, create a simple black square as last resort
            let bmi = Gdi::BITMAPINFO {
                bmiHeader: Gdi::BITMAPINFOHEADER {
                    biSize: std::mem::size_of::<Gdi::BITMAPINFOHEADER>() as u32,
                    biWidth: size as i32,
                    biHeight: -(size as i32), // Negative for top-down DIB
                    biPlanes: 1,
                    biBitCount: 32,
                    biCompression: Gdi::BI_RGB.0 as u32,
                    ..Default::default()
                },
                ..Default::default()
            };

            let mut dib_data: *mut std::ffi::c_void = std::ptr::null_mut();
            let hbitmap_handle: Gdi::HBITMAP = unsafe {
                Gdi::CreateDIBSection(None, &bmi, Gdi::DIB_RGB_COLORS, &mut dib_data, None, 0)
            }?;
            let hbitmap_guard = HBitmapGuard::new(hbitmap_handle);

            if !dib_data.is_null() {
                // Fill with solid black (BGRA format: 0xFF000000)
                let pixel_count = (size * size) as usize;
                let buffer: &mut [u32] = unsafe {
                    std::slice::from_raw_parts_mut(dib_data as *mut u32, pixel_count)
                };

                // Solid black with full alpha
                buffer.fill(0xFF000000);
            }

            log_message("create_fallback_thumbnail: Successfully created bitmap-based fallback");
            Ok(hbitmap_guard.release())
        }
    }
}

// =================================================================
//                      COM Class Factory
// =================================================================

#[implement(Com::IClassFactory)]
struct ClassFactory;

impl Default for ClassFactory {
    fn default() -> Self {
        dll_add_ref();
        log_message("ClassFactory: Created new instance");
        Self {}
    }
}

impl Drop for ClassFactory {
    fn drop(&mut self) {
        log_message("ClassFactory: Dropping instance");
        dll_release();
    }
}

impl Com::IClassFactory_Impl for ClassFactory_Impl {
    #[allow(non_snake_case)]
    fn CreateInstance(&self, punkouter: Ref<'_, IUnknown>, riid: *const GUID, ppvobject: *mut *mut std::ffi::c_void) -> Result<()> {
        ffi_guard!(Result<()>, {
            // log_message(&format!("ClassFactory::CreateInstance: Entered. Requesting interface: {:?}", unsafe { *riid }));

            // Safety checks for null pointers
            if riid.is_null() || ppvobject.is_null() {
                log_message("ClassFactory::CreateInstance: Error - Null pointer passed");
                return Err(Error::new(E_POINTER, "Null pointer passed to CreateInstance"));
            }

            // We do not support aggregation.
            if !punkouter.is_null() {
                log_message("ClassFactory::CreateInstance: Error - Aggregation not supported.");
                return Err(Error::new(CLASS_E_NOAGGREGATION, "Aggregation not supported"));
            }

            log_message("ClassFactory::CreateInstance: Creating ThumbnailProvider instance");

            // Create an instance of our ThumbnailProvider
            let thumbnail_provider: IUnknown = ThumbnailProvider::default().into();

            // Query for the interface requested by the caller and return it.
            let hr: HRESULT = unsafe { thumbnail_provider.query(&*riid, ppvobject) };

            if hr.is_ok() {
                Ok(())
            } else {
                log_message(&format!("ClassFactory::CreateInstance: Error - Exiting with HRESULT: {:?}", hr));
                Err(Error::new(hr, "Failed to query interface"))
            }
        })
    }

    #[allow(non_snake_case)]
    fn LockServer(&self, flock: BOOL) -> Result<()> {
        ffi_guard!(Result<()>, {
            if flock.as_bool() {
                log_message("ClassFactory::LockServer: Locking server (adding reference)");
                dll_add_ref();
            } else {
                log_message("ClassFactory::LockServer: Unlocking server (releasing reference)");
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
// Global flag for whether to enable debug logging
static ENABLE_DEBUG_LOGGING: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);
// A global OnceLock for the log file path, initialized only once
static LOG_FILE_PATH: OnceLock<Option<PathBuf>> = OnceLock::new();

fn dll_add_ref() {
    let new_count = DLL_REFERENCES.fetch_add(1, Ordering::Relaxed) + 1;
    log_message(&format!("DLL reference added. New count: {}", new_count));
}
fn dll_release() {
    let old_count = DLL_REFERENCES.fetch_sub(1, Ordering::Release);
    log_message(&format!("DLL reference released. New count: {}", old_count - 1));
}

/// Generic function to read registry values from HKEY_CLASSES_ROOT\.svg
/// Returns the value as a u32 if it exists and is a valid DWORD, otherwise returns None
fn read_svg_registry_dword(value_name: &str) -> Option<u32> {
    let mut svg_key: HKEY = HKEY::default();
    let result = unsafe {
        RegOpenKeyExW(
            HKEY_CLASSES_ROOT,
            w!(".svg"),
            Some(0),
            KEY_READ,
            &mut svg_key,
        )
    };

    if result.is_ok() {
        let svg_key_guard = RegistryKeyGuard(svg_key);

        let mut value: u32 = 0;
        let mut value_size = std::mem::size_of::<u32>() as u32;
        let mut value_type = REG_DWORD;

        // Convert the value name to a wide string
        let wide_name = to_pcwstr(value_name);

        let query_result = unsafe {
            RegQueryValueExW(
                svg_key_guard.0,
                PCWSTR(wide_name.as_ptr()),
                None,
                Some(&mut value_type),
                Some(&mut value as *mut u32 as *mut u8),
                Some(&mut value_size),
            )
        };

        // Only return the value if it exists, is a DWORD, and has the expected size
        if query_result.is_ok() && value_type == REG_DWORD && value_size == std::mem::size_of::<u32>() as u32 {
            return Some(value);
        } else if !query_result.is_ok() {
            log_message(&format!("Registry read failed for '{}': {:?}", value_name, query_result));
        }
    } // Registry key automatically closed here by RegistryKeyGuard

    return None
}

// Checks registry for setting for whether to enable debug logging
fn check_debug_logging_registry() {
    // Note: We can't log here initially since logging might not be enabled yet
    let enable_debug = match read_svg_registry_dword("win_sdr_thumbs_enable_debug_log") {
        Some(1) => true,  // Only enable debug logging if value exists and equals 1
        _ => false,       // Default to disabled for any other case (missing, 0, or other values)
    };

    ENABLE_DEBUG_LOGGING.store(enable_debug, Ordering::Relaxed);

    // Now we can log since the flag is set
    if enable_debug {
        log_message("Debug logging ENABLED via registry");
    }
}

// This is our thumbnail provider's unique Class ID (CLSID).
// Use a new GUID for your own projects!
const CLSID_SVG_THUMBNAIL_PROVIDER: GUID = GUID::from_u128(0xadfa4c4b_5cfb_4335_be68_d4d60f2ab71f);

#[no_mangle]
#[allow(non_snake_case)]
extern "system" fn DllMain(hinst_dll: HMODULE, fdw_reason: u32, _lpv_reserved: *const std::ffi::c_void) -> BOOL {
    ffi_guard!(BOOL, {
        if fdw_reason == System::SystemServices::DLL_PROCESS_ATTACH {
            MODULE_HANDLE.store(hinst_dll.0 as *mut _, Ordering::Release);
            // Check registry for debug logging preference once at startup
            check_debug_logging_registry();

            log_message("DllMain: DLL_PROCESS_ATTACH completed. DLL is loaded and initialized.");
        } else if fdw_reason == System::SystemServices::DLL_PROCESS_DETACH {
            log_message("DllMain: DLL_PROCESS_DETACH received. DLL is unloading.");
        }
        true
    })
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllGetClassObject(rclsid: *const GUID, riid: *const GUID, ppv: *mut *mut std::ffi::c_void) -> HRESULT {
    ffi_guard!(HRESULT, {
        // Check registry settings at entry point in case they changed since DLL load
        check_debug_logging_registry();

        log_message("DllGetClassObject: Entered");

        // Safety checks for null pointers
        if rclsid.is_null() || riid.is_null() || ppv.is_null() {
            log_message("DllGetClassObject: Error - Null pointer passed");
            return E_POINTER;
        }

        // Check if the caller is asking for our specific class.
        if unsafe { *rclsid } != CLSID_SVG_THUMBNAIL_PROVIDER {
            log_message(&format!("DllGetClassObject: Error - CLSID mismatch. Requested: {:?}, Expected: {:?}", unsafe { *rclsid }, CLSID_SVG_THUMBNAIL_PROVIDER));
            return CLASS_E_CLASSNOTAVAILABLE;
        }

        log_message("DllGetClassObject: Creating class factory for SVG Thumbnail Provider");

        // Create our class factory.
        let factory: Com::IClassFactory = ClassFactory::default().into();

        // Query for the interface the caller wants (usually IClassFactory) and return it.
        let hr: HRESULT = unsafe { factory.query(riid, ppv) };

        // The factory variable will automatically drop here, releasing our local reference.
        // The caller retains their reference from the query() call.

        // log_message(&format!("DllGetClassObject: Exiting with HRESULT: {:?}", hr));
        // Log only if it's an error
        if hr.is_err() {
            log_message(&format!("DllGetClassObject: Error - Exiting with HRESULT: {:?}", hr));
        } else {
            // log_message("DllGetClassObject: Succeeded.");
        }

        hr
    })
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllCanUnloadNow() -> HRESULT {
    ffi_guard!(HRESULT, {
        let ref_count = DLL_REFERENCES.load(Ordering::Acquire);

        if ref_count == 0 {
            log_message("DllCanUnloadNow: Returning S_OK - DLL can be unloaded");
            S_OK
        } else {
            log_message(&format!("DllCanUnloadNow: Returning S_FALSE - DLL still has {} active references", ref_count));
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
    log_message("create_registry_keys: Starting registry key creation");

    let clsid_string = format!("{{{CLSID_SVG_THUMBNAIL_PROVIDER:?}}}");
    let dll_path = get_dll_path()?;
    log_message(&format!("create_registry_keys: Using CLSID: {} and DLL path: {}", clsid_string, dll_path));

    // Create CLSID\{our-clsid}
    // log_message("create_registry_keys: Creating CLSID root key");
    let clsid_root_key = RegistryKeyGuard::create_root_key(HKEY_CLASSES_ROOT, &w!("CLSID"))?;

    log_message("create_registry_keys: Creating CLSID subkey and setting description");
    let clsid_key = clsid_root_key.create_subkey(&PCWSTR(to_pcwstr(&clsid_string).as_ptr()))?;
    clsid_key.set_string_value("", "SVG Thumbnail Provider (Rust)")?;

    // Create CLSID\{our-clsid}\InprocServer32
    log_message("create_registry_keys: Creating InprocServer32 key");
    let inproc_key = clsid_key.create_subkey(&w!("InprocServer32"))?;
    inproc_key.set_string_value("", &dll_path)?;
    inproc_key.set_string_value("ThreadingModel", "Apartment")?;

    // Associate with .svg files
    log_message("create_registry_keys: Associating with .svg files");
    let svg_root_key = RegistryKeyGuard(HKEY_CLASSES_ROOT).create_subkey(&w!(".svg"))?;
    let svg_shellex_key = svg_root_key.create_subkey(&w!("shellex"))?;
    let svg_handler_key = svg_shellex_key.create_subkey(&w!("{E357FCCD-A995-4576-B01F-234630154E96}"))?;
    svg_handler_key.set_string_value("", &clsid_string)?;

    // Associate with .svgz files
    log_message("create_registry_keys: Associating with .svgz files");
    let svgz_root_key = RegistryKeyGuard(HKEY_CLASSES_ROOT).create_subkey(&w!(".svgz"))?;
    let svgz_shellex_key = svgz_root_key.create_subkey(&w!("shellex"))?;
    let svgz_handler_key = svgz_shellex_key.create_subkey(&w!("{E357FCCD-A995-4576-B01F-234630154E96}"))?;
    svgz_handler_key.set_string_value("", &clsid_string)?;

    // log_message("create_registry_keys: Notifying shell of association changes");
    unsafe { Shell::SHChangeNotify(Shell::SHCNE_ASSOCCHANGED, Shell::SHCNF_IDLIST, None, None) };

    // log_message("create_registry_keys: Successfully completed registry key creation");
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
        }
        if key.is_invalid() {
            return Err(Error::new(E_FAIL, "RegCreateKeyExW returned null handle"));
        }

        Ok(RegistryKeyGuard(key))
    }

    // fn get(&self) -> HKEY {
    //     self.0
    // }

    fn create_root_key(hive: HKEY, name: &PCWSTR) -> Result<RegistryKeyGuard> {
        let mut key = HKEY::default();
        unsafe {
            RegCreateKeyExW(
                hive,
                *name,
                None,
                None,
                REG_OPTION_NON_VOLATILE,
                WRITE_FLAGS,
                None,
                &mut key,
                None
            ).ok()?;
        }
        Ok(RegistryKeyGuard(key))
    }

    /// Sets a REG_SZ (string) value for this registry key.
    /// The `name` can be an empty string to set the (Default) value.
    fn set_string_value(&self, name: &str, value: &str) -> Result<()> {
        // Convert name and value to null-terminated wide strings (Vec<u16>)
        let wide_name = to_pcwstr(name);
        let wide_value = to_pcwstr(value);

        // The size for RegSetValueExW must be in bytes, including the null terminator.
        // to_pcwstr already adds the null terminator, so wide_value.len() is correct.
        let value_size_bytes = (wide_value.len() * std::mem::size_of::<u16>()) as u32;

        unsafe {
            RegSetValueExW(
                self.0,
                PCWSTR(wide_name.as_ptr()),
                None,
                REG_SZ,
                Some(std::slice::from_raw_parts(
                    wide_value.as_ptr() as *const u8,
                    value_size_bytes as usize,
                )),
            ).ok()?;
        }
        Ok(())
    }
}

fn delete_registry_keys() -> Result<()> {
    log_message("delete_registry_keys: Starting registry key deletion");

    let clsid_string = format!("{{{CLSID_SVG_THUMBNAIL_PROVIDER:?}}}");
    log_message(&format!("delete_registry_keys: Deleting keys for CLSID: {}", clsid_string));
    // Track if we encountered any real errors (not just "not found")
    let mut first_real_error: Option<Error> = None;

    // Helper closure for robust key deletion
    let mut delete_key_with_error_tracking = |key_path: PCWSTR| {
        let result = unsafe { RegDeleteKeyExW(HKEY_CLASSES_ROOT, key_path, WRITE_FLAGS.0, Some(0)) };
        if result == ERROR_SUCCESS || result == ERROR_FILE_NOT_FOUND {
            // Success or key already gone - both fine for uninstall
        } else {
            // Real error (access denied, etc.) - remember the first one we see
            if first_real_error.is_none() {
                first_real_error = Some(Error::new(result.into(), "Registry key deletion failed"));
            }
        }
    };

    // Try to delete all keys, tracking errors but not stopping
    let inproc_path = to_pcwstr(&format!("CLSID\\{}\\InprocServer32", clsid_string));
    delete_key_with_error_tracking(PCWSTR(inproc_path.as_ptr()));

    let clsid_path = to_pcwstr(&format!("CLSID\\{}", clsid_string));
    delete_key_with_error_tracking(PCWSTR(clsid_path.as_ptr()));

    delete_key_with_error_tracking(w!(".svg\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}"));
    delete_key_with_error_tracking(w!(".svgz\\shellex\\{E357FCCD-A995-4576-B01F-234630154E96}"));

    // Always notify of association changes, even if some deletions failed
    unsafe { Shell::SHChangeNotify(Shell::SHCNE_ASSOCCHANGED, Shell::SHCNF_IDLIST, None, None) };

    // Now propagate the first real error we encountered, if any
    match first_real_error {
        Some(error) => Err(error),
        None => Ok(()),
    }
}


#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllRegisterServer() -> HRESULT {
    ffi_guard!(HRESULT, {
        // log_message("DllRegisterServer: Starting registration");
        match create_registry_keys() {
            Ok(_) => {
                log_message("DllRegisterServer: Registration succeeded");
                S_OK
            },
            Err(e) => {
                log_message(&format!("DllRegisterServer: Registration failed: {:?}", e));
                E_FAIL
            },
        }
    })
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "system" fn DllUnregisterServer() -> HRESULT {
    ffi_guard!(HRESULT, {
        // log_message("DllUnregisterServer: Starting unregistration");
        match delete_registry_keys() {
            Ok(_) => {
                log_message("DllUnregisterServer: Unregistration succeeded");
                S_OK
            },
            Err(e) => {
                log_message(&format!("DllUnregisterServer: Unregistration failed: {:?}", e));
                E_FAIL
            },
        }
    })
}

#[no_mangle]
// Simple function that only notifies the shell of file association changes.
pub extern "system" fn notify_shell_change() -> HRESULT {
    ffi_guard!(HRESULT, {
        // Notify the shell that file associations have changed
        unsafe { Shell::SHChangeNotify(Shell::SHCNE_ASSOCCHANGED, Shell::SHCNF_IDLIST, None, None) };
        S_OK
    })
}

// =================================================================

// -------------- Logger ----------------
fn log_message(message: &str) {
    if !ENABLE_DEBUG_LOGGING.load(Ordering::Relaxed) {
        return;
    }

    // get_or_init will only execute the closure ONCE, the very first time it's called.
    // All subsequent calls will return the cached value instantly.
    let log_path_option = LOG_FILE_PATH.get_or_init(|| {
        let known_folder_flags = Shell::KNOWN_FOLDER_FLAG::default(); // Use default flags, no special options
        let desktop_path_pwstr = match unsafe { SHGetKnownFolderPath(&FOLDERID_Desktop, known_folder_flags, None) } {
            Ok(path) => path,
            Err(_) => return None, // Initialization failed, cache 'None'
        };

        let desktop_path_guard = CoTaskMemFreeGuard(desktop_path_pwstr);

        let mut path = match unsafe { desktop_path_guard.0.to_string() } {
            Ok(s) => PathBuf::from(s),
            Err(_) => return None, // Conversion failed, cache 'None'
        };

        path.push("win_sdr_thumbs_debug_log.txt");
        Some(path) // Success! Cache the full path.
        // --- End of one-time execution block ---
    });

    // Now, use the cached path.
    // If initialization failed, log_path_option will be &None, and we'll do nothing.
    if let Some(log_path) = log_path_option {
        match std::fs::OpenOptions::new().create(true).append(true).open(log_path) {
            Ok(mut file) => {
                let pid = std::process::id();
                let tid = std::thread::current().id();
                let time_str = get_formatted_time_string_win_api();

                let _ = writeln!(file, "[PID: {} | TID: {:?}] [{}] {}", pid, tid, time_str, message);
            }
            Err(_) => {
                // Opening the file failed.
            }
        }
    }
}

fn get_formatted_time_string_win_api() -> String {
    let system_time = unsafe { GetLocalTime() };
    let mut time_buffer = [0u16; 64]; // Buffer for formatted time string
    let chars_written = unsafe { // Returns the number of characters put into the buffer. If 0 it failed
        GetTimeFormatEx(
            None, // Null is LOCALE_NAME_USER_DEFAULT
            TIME_FORMAT_FLAGS::default(), // Default time format flags
            Some(&system_time),
            None, // Use default format
            Some(&mut time_buffer),
        )
    };

    if chars_written > 0 {
        return String::from_utf16_lossy(&time_buffer[..chars_written as usize - 1]) // -1 to remove null terminator
    } else {
        // Fallback if formatting fails
         return format!("{:02}:{:02}:{:02}.{:03}",
            system_time.wHour,
            system_time.wMinute,
            system_time.wSecond,
            system_time.wMilliseconds)
    }
}

// fn log_message(message: &str) {
//     println!("{}", message);
// }
