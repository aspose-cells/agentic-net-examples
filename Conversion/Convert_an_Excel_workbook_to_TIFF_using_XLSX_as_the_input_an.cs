using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConvertExcelToTiff
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure image options for TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.SaveFormat = SaveFormat.Tiff; // Use the rule-provided property

        // Create a sheet renderer for the first worksheet
        SheetRender renderer = new SheetRender(workbook.Worksheets[0], options);

        // Render the worksheet to a multi‑page TIFF file
        renderer.ToTiff("output.tiff");

        Console.WriteLine("Excel workbook successfully converted to TIFF.");
    }
}