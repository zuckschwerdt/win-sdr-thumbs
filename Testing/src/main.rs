use std::fs::File;
use std::io::Read;
use std::ptr;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        Graphics::Gdi::*,
        UI::WindowsAndMessaging::*,
        System::LibraryLoader::GetModuleHandleW,
    },
};

use win_svg_thumbs::render_svg_to_hbitmap;

// Global variable to store the HBITMAP so it can be accessed in the window procedure
static mut GLOBAL_HBITMAP: HBITMAP = HBITMAP(ptr::null_mut());

fn main() -> Result<()> {
    // Check if the SVG file exists before proceeding
    let mut svg_path = std::env::current_dir().expect("Failed to get current directory");
    svg_path.push("test.svg");
    if !svg_path.exists() {
        use std::os::windows::ffi::OsStrExt;
        let full_path = svg_path.canonicalize().unwrap_or(svg_path.clone());
        let full_path_str = full_path.display().to_string();
        let msg = format!("Could not find SVG file at: {}", full_path_str);
        // Convert Rust String to wide string for MessageBoxW
        let wide: Vec<u16> = std::ffi::OsStr::new(&msg).encode_wide().chain(std::iter::once(0)).collect();
        unsafe {
            MessageBoxW(
                None,
                PCWSTR(wide.as_ptr()),
                w!("File Not Found"),
                MB_OK | MB_ICONERROR,
            );
        }
        return Ok(());
    }

    // Load SVG data from a file (test.svg)
    let mut file = File::open(&svg_path).expect("Failed to open SVG file");
    let mut svg_data = Vec::new();
    file.read_to_end(&mut svg_data).expect("Failed to read SVG file");

    // Set desired output size
    let width = 256;
    let height = 256;

    // Render SVG to HBITMAP
    let hbitmap = render_svg_to_hbitmap(&svg_data, width, height)?;
    println!("Successfully rendered SVG to HBITMAP: {:?}", hbitmap);

    // Store the HBITMAP globally so the window procedure can access it
    unsafe {
        GLOBAL_HBITMAP = hbitmap;
    }

    // Create a window to display the image
    unsafe {
        let h_instance = GetModuleHandleW(None)?;
        let class_name = w!("SvgImageWindow");        // Register window class
        let wc = WNDCLASSW {
            hInstance: HINSTANCE(h_instance.0),
            lpszClassName: class_name,
            lpfnWndProc: Some(window_proc),
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as *mut _),
            ..Default::default()
        };
        
        if RegisterClassW(&wc) == 0 {
            return Err(Error::from_win32());
        }

        // Create the window
        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            class_name,
            w!("SVG Image"),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            width as i32 + 50, // Add some padding for window frame
            height as i32 + 70, // Add some padding for window frame and title bar
            None,
            None,
            Some(HINSTANCE(h_instance.0)),
            None,
        )?;

        let _ = ShowWindow(hwnd, SW_SHOW);

        // Message loop
        let mut msg = MSG::default();
        while GetMessageW(&mut msg, None, 0, 0).into() {
            let _ = TranslateMessage(&msg);
            let _ = DispatchMessageW(&msg);
        }

        // Clean up the HBITMAP resource
        let hbitmap: HBITMAP = GLOBAL_HBITMAP;
        if !hbitmap.is_invalid() {
            let _ = DeleteObject(HGDIOBJ(hbitmap.0));
        }
    }

    Ok(())
}

extern "system" fn window_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_PAINT => {
            unsafe {
                let mut ps = PAINTSTRUCT::default();
                let hdc = BeginPaint(hwnd, &mut ps);

                let hbitmap = GLOBAL_HBITMAP;
                if !hbitmap.is_invalid() {
                    // Create a compatible device context
                    let hdc_mem = CreateCompatibleDC(Some(hdc));
                    // Select the bitmap into the memory DC
                    let old_bitmap = SelectObject(hdc_mem, HGDIOBJ(hbitmap.0));

                    // Get the bitmap dimensions
                    let mut bitmap = BITMAP::default();
                    GetObjectW(
                        HGDIOBJ(hbitmap.0),
                        std::mem::size_of::<BITMAP>() as i32,
                        Some(&mut bitmap as *mut _ as *mut _)
                    );

                    // Use AlphaBlend to respect the alpha channel
                    let blend_func = BLENDFUNCTION {
                        BlendOp: AC_SRC_OVER as u8,
                        BlendFlags: 0,
                        SourceConstantAlpha: 255,
                        AlphaFormat: AC_SRC_ALPHA as u8,
                    };
                    let _ = AlphaBlend(
                        hdc,
                        10, 10, // Position in window
                        bitmap.bmWidth,
                        bitmap.bmHeight,
                        hdc_mem,
                        0, 0, // Source position
                        bitmap.bmWidth,
                        bitmap.bmHeight,
                        blend_func,
                    );

                    // Clean up
                    SelectObject(hdc_mem, old_bitmap);
                    let _ = DeleteDC(hdc_mem);
                }

                let _ = EndPaint(hwnd, &ps);
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            unsafe { PostQuitMessage(0) };
            LRESULT(0)
        }
        _ => unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) },
    }
}