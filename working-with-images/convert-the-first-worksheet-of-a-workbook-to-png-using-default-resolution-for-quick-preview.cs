// Title: Create a PNG preview of the first worksheet in an Excel file using Aspose.Cells for .NET with default resolution
// AI Prompts: Write C# code that loads an .xlsx workbook and saves the first worksheet as a PNG image using Aspose.Cells with default ImageOrPrintOptions. | Show how to use SheetRender to render only the first page of a worksheet to a PNG file without setting a custom resolution. | Provide a minimal example that converts the first sheet of an Excel workbook to a PNG preview using Aspose.Cells for .NET.
// Common Searches: asp.net render first worksheet to png using aspose.cells default settings | c# generate png preview of excel sheet with aspose.cells without specifying resolution | how to export only the first page of an Excel worksheet as png in .net | aspose.cells quick png snapshot of first sheet from xlsx file
// Tags: SheetRender ToImage PNG export | default ImageOrPrintOptions rendering | export first worksheet as PNG Aspose.Cells | quick Excel preview image .NET | render worksheet page to image

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads 'input.xlsx', renders the first worksheet with default image options, and saves page 0 as 'output.png' using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Use default image options (default resolution)
        ImageOrPrintOptions options = new ImageOrPrintOptions();

        // Create a renderer for the worksheet
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Render the first page of the worksheet to a PNG file
        sheetRender.ToImage(0, "output.png");
    }
}
