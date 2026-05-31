using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Multi‑page TIFF rendering demo");
        // Add more rows to ensure the workbook spans multiple pages
        sheet.Cells["A1000"].PutValue("End of data");

        // Configure default image options for TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Tiff; // Default TIFF format

        // Create a workbook renderer with the specified options
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the entire workbook to a multi‑page TIFF file
        renderer.ToImage("output.tiff");

        Console.WriteLine("Workbook successfully rendered to multi‑page TIFF: output.tiff");
    }
}