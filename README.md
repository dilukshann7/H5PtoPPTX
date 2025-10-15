# H5PtoPPTX

H5PtoPPTX is a Windows desktop application that converts H5P interactive content into PowerPoint presentations. This tool extracts images from `.h5p` files and generates a `.pptx` presentation with each image on a separate slide.

## Features

*   **Bulk Conversion:** Convert multiple `.h5p` files in a single batch process.
*   **Image Extraction:** Automatically extracts images from your H5P content.
*   **PowerPoint Generation:** Creates a `.pptx` file for each `.h5p` file, with one image per slide.
*   **User-Friendly Interface:** Simple and intuitive graphical user interface.
*   **Logging:** View the status of each conversion in the log window.

## How to Use

1.  **Download the Application:**
    *   Go to the [Releases](https://github.com/DilukshanN7/H5PtoPPTX/releases) page.
    *   Download the latest `H5PtoPPTX.zip` file.
    *   Extract the contents of the zip file to a folder on your computer.

2.  **Run the Application:**
    *   Double-click on `H5PtoPPTX.exe` to launch the application.

3.  **Convert Your Files:**
    *   Click the "Browse" button to select the input folder containing your `.h5p` files.
    *   Click the "Browse" button to select the output folder where you want to save the converted `.pptx` files.
    *   Click the "Convert" button to start the conversion process.

## Building from Source

If you want to build the application from source, you will need:

*   Visual Studio 2022 or later
*   .NET Framework 4.7.2 or later

Clone the repository and open the `H5PtoPPTX.sln` file in Visual Studio.

## Dependencies

This project uses the following libraries:

*   [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.