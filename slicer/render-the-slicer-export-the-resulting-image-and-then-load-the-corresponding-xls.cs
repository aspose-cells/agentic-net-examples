using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace SlicerRenderExample
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source workbook that already contains a slicer
            string sourceWorkbookPath = "input.xlsx";

            // Load the workbook (creation rule: Workbook(string))
            Workbook workbook = new Workbook(sourceWorkbookPath);

            // Get the first worksheet (assumed to contain the slicer)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image rendering options (ImageOrPrintOptions)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png, // render as PNG
                OnePagePerSheet = true                         // render each sheet as a single page
            };

            // Create a SheetRender instance (constructor rule)
            SheetRender sheetRender = new SheetRender(worksheet, renderOptions);

            // Export the rendered image of the worksheet (including the slicer) to a file
            string imageOutputPath = "slicer_render.png";
            sheetRender.ToImage(0, imageOutputPath); // ToImage(int, string) rule

            // Optionally dispose the renderer
            sheetRender.Dispose();

            // Save the original workbook to a new file (save rule)
            string savedWorkbookPath = "output.xlsx";
            workbook.Save(savedWorkbookPath);

            // Load the saved workbook (load rule)
            Workbook loadedWorkbook = new Workbook(savedWorkbookPath);

            // At this point, the slicer image has been exported and the workbook reloaded.
            Console.WriteLine($"Slicer rendered to image: {imageOutputPath}");
            Console.WriteLine($"Workbook saved to: {savedWorkbookPath}");
            Console.WriteLine($"Workbook reloaded successfully.");
        }
    }
}