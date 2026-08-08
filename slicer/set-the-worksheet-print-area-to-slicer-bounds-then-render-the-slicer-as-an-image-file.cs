// Title: Set Print Area to Slicer Bounds and Export Slicer as PNG with Aspose.Cells for .NET
// Description: This C# example loads a workbook, reads the first slicer's shape coordinates, sets the worksheet's print area to those cells, and uses ImageOrPrintOptions (OnlyArea=true) to render the slicer to a PNG file via SheetRender.
// Keywords: Aspose.Cells | C# | slicer | print area | render slicer image | ImageOrPrintOptions | OnlyArea | SheetRender | export slicer PNG | Excel slicer bounds
// Common Searches: Aspose.Cells set print area to slicer | export slicer as PNG C# | render slicer only area Aspose.Cells | get slicer shape coordinates .NET | save slicer image from Excel
// Developer Intent: Set the worksheet's print area to match a slicer's dimensions and generate an image of that slicer.
// Use Cases: Create thumbnail images of slicers for dashboard reports | Automate snapshot generation of slicers for documentation | Provide slicer visuals on web portals or mobile apps | Batch export multiple slicers for reporting pipelines
// AI Prompts: How can I export the slicer as a JPEG instead of PNG using Aspose.Cells? | Show code to render the slicer to a MemoryStream for an ASP.NET Core API response. | Provide a loop that processes all slicers in a worksheet and saves each as a separate image file. | Explain how to adjust DPI or image dimensions when rendering slicer images.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

// This C# example loads a workbook, reads the first slicer's shape coordinates, sets the worksheet's print area to those cells, and uses ImageOrPrintOptions (OnlyArea=true) to render the slicer to a PNG file via SheetRender.
class RenderSlicerExample
{
    static void Main()
    {
        // Load an existing workbook that contains a slicer
        // (Replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one slicer
        if (sheet.Slicers.Count == 0)
        {
            Console.WriteLine("No slicer found in the worksheet.");
            return;
        }

        // Get the first slicer
        Slicer slicer = sheet.Slicers[0];

        // Obtain the underlying shape to read its bounds
        SlicerShape shape = slicer.Shape;

        // Upper‑left cell of the slicer
        int startRow = shape.UpperLeftRow;
        int startCol = shape.UpperLeftColumn;

        // Lower‑right cell of the slicer
        int endRow = shape.LowerRightRow;
        int endCol = shape.LowerRightColumn;

        // Convert cell indexes to A1 style names
        string startCell = CellsHelper.CellIndexToName(startRow, startCol);
        string endCell   = CellsHelper.CellIndexToName(endRow, endCol);

        // Set the worksheet print area to exactly the slicer bounds
        sheet.PageSetup.PrintArea = $"{startCell}:{endCell}";

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;   // Output format
        options.OnlyArea = true;            // Render only the defined print area

        // Create a SheetRender for the worksheet with the above options
        SheetRender render = new SheetRender(sheet, options);

        // Render the first (and only) page to an image file
        string outputPath = "slicer.png";
        render.ToImage(0, outputPath);

        Console.WriteLine($"Slicer rendered to image: {outputPath}");
    }
}
