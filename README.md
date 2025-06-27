# Windows SVG Thumbnail Provider in Rust

A high-performance thumbnail provider for Windows that generates explorer thumbnails for `.svg` and `.svgz` files, written in Rust.

## Screenshot
<p align="center">
<img width=650 src=https://github.com/user-attachments/assets/67050436-809e-437b-9c17-4cdeeb386450>
</p>

## How to Install

1.  Go to the [Releases](https://github.com/ThioJoe/win-svg-thumbs-rust/releases) page.
2.  Download the latest `win_svg_thumbs.dll` file. Put it somewhere it can remain.
3.  Open a Command Prompt with administrator privileges.
4.  Navigate to the directory where you saved the `.dll` file.
5.  Run the following command to register the DLL:
    ```
    regsvr32 win_svg_thumbs.dll
    ```

## Usage

Once the DLL is registered, Windows Explorer will automatically use this provider to display thumbnails for `.svg` and `.svgz` files.

To uninstall, run the following command in an administrator Command Prompt:
  ```
  regsvr32 /u win_svg_thumbs.dll
  ```

## How to Compile it Yourself

### Prerequisites

* Setup Rust on Windows: [See these instructions](https://learn.microsoft.com/en-us/windows/dev-environment/rust/setup)

### Instructions

1.  Clone the repository:
    ```
    git clone https://github.com/ThioJoe/win-svg-thumbs-rust
    ```
2.  Navigate to the project directory:
    ```
    cd win-svg-thumbs-rust
    ```
3.  Build the project in release mode:
    ```
    cargo build --release
    ```
4.  The compiled DLL will be located in the `target/release` directory.
