using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class RenderSlicerExample
{
    static void Main()
    {
        // Path to the source Excel file that contains a slicer
        string inputPath = "input.xlsx";

        // Load the workbook from the file (lifecycle: load)
        Workbook workbook = new Workbook(inputPath);

        // Access the worksheet that holds the slicer (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,   // Output format
            OnePagePerSheet = true       // Render the whole sheet as a single page
        };

        // Create a SheetRender instance for the worksheet (lifecycle: create)
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Render the first (and only) page of the sheet to an image file
        // This will include the slicer shape in the rendered output
        string outputImagePath = "slicer_render.png";
        sheetRender.ToImage(0, outputImagePath); // lifecycle: save (image file)

        Console.WriteLine($"Worksheet with slicer rendered successfully to: {outputImagePath}");
    }
}