using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

class RenderSlicerImage
{
    static void Main()
    {
        // Load the existing workbook (replace with actual path if needed)
        Workbook workbook = new Workbook("sample.xlsx");

        // Access the first worksheet (adjust index if slicer is on another sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one slicer
        if (sheet.Slicers.Count == 0)
        {
            Console.WriteLine("No slicer found in the worksheet.");
            return;
        }

        // Retrieve the first slicer
        Slicer slicer = sheet.Slicers[0];

        // Access the associated shape (optional, demonstrates locating the slicer)
        Shape slicerShape = slicer.Shape;
        Console.WriteLine($"Slicer Shape Name: {slicerShape.Name}");

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            OnePagePerSheet = true
        };

        // Create a renderer for the workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the first page (which contains the slicer) to an image file
        renderer.ToImage(0, "slicer.png");

        Console.WriteLine("Slicer rendered and saved as slicer.png");
    }
}