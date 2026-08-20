// Title: Render the First Worksheet to PNG with Aspose.Cells for .NET (default DPI)
// Description: Loads an Excel workbook, configures ImageOrPrintOptions for PNG using the library's default resolution, and uses SheetRender to export the first worksheet as a single‑page PNG file (first_worksheet_preview.png).
// Keywords: Aspose.Cells | C# | SheetRender | PNG preview | default DPI | export worksheet as image | ImageOrPrintOptions | Excel to PNG | quick preview
// Common Searches: Aspose.Cells render first sheet to PNG | C# export Excel worksheet as PNG default resolution | How to create worksheet thumbnail with Aspose.Cells | Generate PNG preview of Excel file using .NET | SheetRender PNG output example
// Developer Intent: Create a PNG snapshot of the first worksheet using Aspose.Cells' default DPI for a fast visual preview.
// Use Cases: Display thumbnail previews of uploaded Excel files in a web portal. | Generate a snapshot for email or PDF reports without altering the original workbook. | Provide an instant visual preview in a desktop application that reads Excel data. | Create image assets for documentation or training materials.
// AI Prompts: Show how to set a custom DPI (e.g., 300) when rendering the first worksheet to PNG. | Explain how to batch‑convert all worksheets in a workbook to separate PNG files. | Provide code for robust error handling when the input file is missing or the worksheet index is out of range. | Demonstrate rendering a specific cell range to PNG instead of the whole sheet. | Show how to write the PNG image to a memory stream for further processing.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPreview
{
    // Loads an Excel workbook, configures ImageOrPrintOptions for PNG using the library's default resolution, and uses SheetRender to export the first worksheet as a single‑page PNG file (first_worksheet_preview.png).
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            string sourcePath = "input.xlsx";
            Workbook workbook = new Workbook(sourcePath);

            // Configure image rendering options for PNG output.
            // No resolution is set, so default resolution will be used.
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                // OnePagePerSheet ensures the whole worksheet is rendered to a single image.
                OnePagePerSheet = true
            };

            // Create a SheetRender for the first worksheet.
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], options);

            // Render the first page (index 0) of the worksheet to a PNG file.
            string outputPath = "first_worksheet_preview.png";
            sheetRender.ToImage(0, outputPath);

            // Clean up resources.
            sheetRender.Dispose();

            Console.WriteLine($"First worksheet rendered to PNG at: {outputPath}");
        }
    }
}
