using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

class RenderSlicerDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Create image/print options for rendering
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;          // Output format
        options.OnePagePerSheet = true;             // Render each sheet as a single page

        // Render the worksheet to an image
        SheetRender render = new SheetRender(sheet, options);

        // Save each rendered page as a separate PNG file
        for (int i = 0; i < render.PageCount; i++)
        {
            render.ToImage(i, $"output_page_{i}.png");
        }

        // Optional: display information about the first slicer on the sheet
        if (sheet.Slicers.Count > 0)
        {
            Slicer slicer = sheet.Slicers[0];
            SlicerShape shape = slicer.Shape;
            Console.WriteLine($"Slicer Shape Name: {shape.Name}");
            Console.WriteLine($"Slicer Shape Width: {shape.Width}");
            Console.WriteLine($"Slicer Shape Height: {shape.Height}");
        }
    }
}