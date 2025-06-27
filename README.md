# Windows SVG Thumbnail Provider in Rust

A high-performance thumbnail provider for Windows that generates explorer thumbnails for `.svg` and `.svgz` files, written in Rust, with no third party dependencies.

## Why Use This One?

<table>
  <tbody>
    <tr>
      <td><b>No Third-Party Dependencies</b></td>
      <td>Built using only official Microsoft-published Rust crates (found in the <a href="https://github.com/microsoft/windows-rs"><code>windows&#8209;rs</code></a> repo).</td>
    </tr>
    <tr>
      <td><b>Works Purely Through Windows API</b></td>
      <td>
        Operates purely through the built-in Windows API for rendering and COM integration. Just a single native DLL.
        <ul>
          ↳ No third party libraries that may go out of date or add overhead
        </ul>
      </td>
    </tr>
    <tr>
      <td><b>Trusted Certificate Signed</b></td>
      <td>
        Signed via Azure Trusted Signing, which requires rigorous verification of real-world identity.
        <ul>
          ↳ Proves it's not made by an unaccountable/unknown developer
        </ul>
      </td>
    </tr>
  </tbody>
</table>

## Screenshot
<p align="center">
<img width=650 src=https://github.com/user-attachments/assets/67050436-809e-437b-9c17-4cdeeb386450>
</p>

## How to Install

1.  Go to the [Releases](https://github.com/ThioJoe/win-svg-thumbs-rust/releases) page.
2.  Download the latest `win_svg_thumbs.dll` file. Put it somewhere it can remain.
3.  Open a Command Prompt with **administrator privileges**.
     - (Administrator is required or you will get error `0x80004005` for lack of permission)
5.  Navigate to the directory where you saved the `.dll` file.
6.  Run the following command to register the DLL:
    ```
    regsvr32 win_svg_thumbs.dll
    ```

## Usage

Once the DLL is registered, Windows Explorer will automatically use this provider to display thumbnails for `.svg` and `.svgz` files.

To uninstall, run the following command in an administrator Command Prompt:
  ```
  regsvr32 /u win_svg_thumbs.dll
  ```

## How it works (Technical Details)

When Windows Explorer needs a thumbnail for a `.svg` file, it interacts with this DLL through a series of steps:

1.  **Initialization**: Explorer provides the `.svg` file's data as a stream to the DLL. The provider reads this entire stream into memory.
2.  **Direct2D Rendering**:
    * The provider uses the **Direct2D** graphics API for high-performance, hardware-accelerated rendering.
    * It creates a GPU-based bitmap to serve as a render target.
    * The Direct2D API turns the SVG data into a `SvgDocument` object. The `viewport` attribute is set to the thumbnail size requested by Explorer.
    * The `width` and `height` attributes are removed from the root `<svg>` element before drawing, which I discovered causes the `DrawSvgDocument` method to autoscale the image to the viewport, avoiding the need to do any manual scaling to fill the thumbnail.
    * The `SvgDocument` is then drawn onto the render target bitmap.
3.  **Pixel Data Transfer**:
    * The rendered image is copied from the GPU render target to a "staging" bitmap on the CPU, which allows the program to access the raw pixel data.
    * This data is in a 32-bit BGRA format with **premultiplied alpha**.
4.  **Creating the Final Thumbnail**:
    * A standard Windows GDI `HBITMAP` is created, which is the final format Explorer needs for the thumbnail.
    * The pixel data is copied from the staging bitmap to the final `HBITMAP`.
       * The GDI requires straight-alpha (as opposed to pre-multiplied), so during this process, the alpha is **un-premultiplied**.
       * The un-premultiplication step is necessary for displaying transparency correctly and prevents dark edges from appearing on the final thumbnail.
5.  **Safety**: The entire thumbnail generation process is wrapped in a panic handler (`catch_unwind`). This ensures that if any unexpected error occurs during rendering, it will not crash the host application (e.g., `explorer.exe`).

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
