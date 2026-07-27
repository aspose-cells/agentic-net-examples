using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class WorksheetToPng
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Optional: add some data to the sheet
        sheet.Cells["A1"].PutValue("Sample Data");

        // Configure image rendering options for PNG (default resolution)
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;

        // Create a SheetRender instance for the worksheet
        SheetRender renderer = new SheetRender(sheet, options);

        // Define the output PNG file path
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FirstSheet.png");

        // Render the first page (index 0) of the worksheet to the PNG file
        renderer.ToImage(0, outputPath);

        Console.WriteLine($"First worksheet rendered to PNG: {outputPath}");
    }
}