using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class RenderWorksheetToImage
{
    static void Main()
    {
        // Load the source Excel file (ensure the file exists at this path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;          // Output format
        options.OnePagePerSheet = true;             // Render the whole sheet on one page

        // Create a SheetRender instance for the worksheet
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Optional: display total page count
        Console.WriteLine("Page count: " + sheetRender.PageCount);

        // Render the first page (index 0) to an image file
        string outputPath = "output.png";
        sheetRender.ToImage(0, outputPath);

        Console.WriteLine("Worksheet rendered to image: " + outputPath);
    }
}