// Title: Convert an Excel workbook to a multi‑page TIFF using Aspose.Cells for .NET (default options)
// Description: Loads an .xlsx file with Aspose.Cells, sets ImageOrPrintOptions.ImageType to TIFF while keeping all other settings at their defaults, creates a WorkbookRender instance for the whole workbook, and writes every worksheet page to a single multi‑page TIFF file.
// Keywords: Aspose.Cells C# | Excel to TIFF conversion | multi‑page TIFF export | WorkbookRender ToImage | ImageOrPrintOptions default | .NET Excel image rendering | export workbook as TIFF | global Excel to image | Aspose.Cells TIFF compression
// Common Searches: Aspose.Cells export entire workbook to TIFF | C# create multi‑page TIFF from Excel | default rendering options TIFF Aspose.Cells | how to render all worksheets to one TIFF file | convert Excel to multi‑page image .NET
// Developer Intent: Generate a single multi‑page TIFF that contains every printed page of an Excel workbook without customizing rendering parameters.
// Use Cases: Archive financial statements as a compact image file for regulatory compliance. | Bundle multiple report sheets into one attachment when PDF is prohibited. | Automate nightly conversion of analytical workbooks to TIFF for secure storage.
// AI Prompts: Provide C# code that converts an Excel workbook to a multi‑page TIFF with a specific DPI while preserving default options. | Explain how to enable LZW compression for TIFF output in Aspose.Cells ImageOrPrintOptions. | Show how to select only certain worksheets for TIFF conversion and keep default settings for the rest.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Loads an .xlsx file with Aspose.Cells, sets ImageOrPrintOptions.ImageType to TIFF while keeping all other settings at their defaults, creates a WorkbookRender instance for the whole workbook, and writes every worksheet page to a single multi‑page TIFF file.
class ConvertWorkbookToMultiPageTiff
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx"); // specify the path to the source file

        // Use default rendering options, only set the image type to TIFF
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Tiff; // ensures TIFF output while keeping other defaults

        // Create a renderer for the entire workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render all pages of the workbook to a multi‑page TIFF file
        renderer.ToImage("output.tiff"); // specify the desired output file name

        Console.WriteLine("Workbook successfully converted to a multi‑page TIFF.");
    }
}
