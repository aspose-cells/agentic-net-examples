// Title: Insert a rectangle shape into an Excel worksheet and assign a unique name using Aspose.Cells for .NET
// AI Prompts: Insert a rectangle graphic at row 2, column 2 with 5‑pixel top/left offsets, assign its Name as 'UniqueShape001', set placement to free‑floating, and save the workbook as an Xlsx file. | Create a shape, set line weight to 1.5 points and dash style to solid, then persist the workbook containing the shape.
// Common Searches: c# aspose.cells insert rectangle shape at specific cell coordinates | aspose.cells assign custom identifier to a shape in an Excel file | configure shape to float above cells using Aspose.Cells | save workbook with added shape as .xlsx using Aspose.Cells
// Tags: add rectangle shape Aspose.Cells C# | set shape Name property Aspose.Cells | set shape placement mode freefloating Aspose.Cells | customize shape line appearance Aspose.Cells | save workbook with shape Aspose.Cells Xlsx

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The program creates a new workbook, inserts a rectangle shape at row 2, column 2 with pixel offsets, assigns it the unique name "UniqueShape001", configures free‑floating placement, customizes line weight and dash style, and saves the file as ShapeDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define shape position and size
        int upperLeftRow = 2;      // Row index (0‑based)
        int upperLeftColumn = 2;   // Column index (0‑based)
        int top = 5;               // Pixels from the top of the cell
        int left = 5;              // Pixels from the left of the cell
        int height = 100;          // Height in pixels
        int width = 200;           // Width in pixels

        // Insert a rectangle shape
        Shape shape = sheet.Shapes.AddShape(
            MsoDrawingType.Rectangle,
            upperLeftRow,
            upperLeftColumn,
            top,
            left,
            height,
            width);

        // Assign a unique name to the shape for later reference
        shape.Name = "UniqueShape001";

        // Optional: set visual properties
        shape.Placement = PlacementType.FreeFloating;
        shape.Line.Weight = 1.5;
        shape.Line.DashStyle = MsoLineDashStyle.Solid;

        // Save the workbook
        workbook.Save("ShapeDemo.xlsx", SaveFormat.Xlsx);
    }
}
