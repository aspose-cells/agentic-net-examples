using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet and its slicers
        foreach (Worksheet ws in workbook.Worksheets)
        {
            SlicerCollection slicers = ws.Slicers;
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];
                // Access the associated shape to obtain visual properties
                var shape = slicer.Shape;
                Console.WriteLine($"Worksheet: {ws.Name}, Slicer: {slicer.Name}, Shape: {shape.Name}, Width: {shape.Width}, Height: {shape.Height}");
            }
        }

        // Set rendering options (PNG images, one page per sheet)
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            OnePagePerSheet = true
        };

        // Create a renderer for the workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render each sheet (including slicers) to an image file
        for (int page = 0; page < renderer.PageCount; page++)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                renderer.ToImage(page, ms);
                File.WriteAllBytes($"Sheet_{page + 1}.png", ms.ToArray());
            }
        }
    }
}