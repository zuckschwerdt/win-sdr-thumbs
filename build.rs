use std::env;
use std::path::{Path, PathBuf};

fn main() {
    let crate_dir = env::var("CARGO_MANIFEST_DIR").unwrap();
    let build_target = env::var("TARGET").unwrap();
    let lib_dir = Path::new(&crate_dir).join("lib").join(build_target);
    let lib_dir = lib_dir.to_str().unwrap();
    println!("cargo:rustc-link-search={lib_dir}");

    let out_dir = env::var("OUT_DIR").unwrap();

    // Detect target architecture
    let target_arch = env::var("CARGO_CFG_TARGET_ARCH").unwrap();
    // let target_triple = env::var("TARGET").unwrap();

    // Calculate the target directory path
    let binding = PathBuf::from(&out_dir);
    let target_dir = binding
        .ancestors()
        .nth(3) // Go up to target/{target}/{profile}
        .unwrap();

    match target_arch.as_str() {
        "x86_64" => {
            let dll_path = target_dir.join("win_sdr_thumbs_x64.dll");
            println!("cargo:rustc-link-arg=/OUT:{}", dll_path.display());
        },
        "x86" => {
            let dll_path = target_dir.join("win_sdr_thumbs_x86.dll");
            println!("cargo:rustc-link-arg=/OUT:{}", dll_path.display());
        },
        "aarch64" => {
            let dll_path = target_dir.join("win_sdr_thumbs_arm64.dll");
            println!("cargo:rustc-link-arg=/OUT:{}", dll_path.display());
        },
        _ => {}
    }

    // println!("cargo:warning=Target arch: {}", target_arch);
    // println!("cargo:warning=Target directory: {}", target_dir.display());
}
