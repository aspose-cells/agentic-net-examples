using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class WorksheetToPng
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        var workbook = new Workbook("input.xlsx");

        // Get the first worksheet (index 0)
        var worksheet = workbook.Worksheets[0];

        // Set image rendering options – PNG format, default resolution
        var options = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png
        };

        // Create a SheetRender for the worksheet with the specified options
        var sheetRender = new SheetRender(worksheet, options);

        // Render the first page (page index 0) to a PNG file
        sheetRender.ToImage(0, "first_sheet.png");

        // Release resources used by SheetRender
        sheetRender.Dispose();

        Console.WriteLine("First worksheet rendered to PNG successfully.");
    }
}