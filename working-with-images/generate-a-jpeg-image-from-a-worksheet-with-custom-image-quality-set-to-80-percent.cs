using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data to the worksheet
        worksheet.Cells["A1"].PutValue("Sample text for JPEG rendering");

        // Configure image options: set JPEG format and custom quality (80%)
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;   // Output format
        options.Quality = 80;                // JPEG quality (0-100)

        // Create a SheetRender instance with the worksheet and options
        SheetRender renderer = new SheetRender(worksheet, options);

        // Render the first page (index 0) to a JPEG file
        renderer.ToImage(0, "WorksheetImage_Quality80.jpg");

        Console.WriteLine("Worksheet rendered to JPEG with quality 80%.");
    }
}