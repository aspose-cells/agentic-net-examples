// Title: Anchor a rectangle shape to a SEQUENCE‑generated dynamic array range and keep it moving and resizing with cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link rectangle shape to SEQUENCE formula result range | move and resize shape with dynamic array cells Aspose.Cells .NET | anchor shape to a formula‑generated range in Excel using Aspose.Cells | update shape position when SEQUENCE array expands in Aspose.Cells | set shape placement to MoveAndSize for dynamic array in C#
// Tags: shape placement MoveAndSize for dynamic arrays | AddShape with SEQUENCE generated range | Aspose.Cells rectangle anchored to formula output | C# dynamic array shape synchronization | Excel shape association with SEQUENCE output using Aspose

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, inserts a SEQUENCE formula in A1 to produce a 5×3 dynamic array, forces calculation, then adds a rectangle shape anchored to the resulting range via AddShape. The shape's Placement is set to MoveAndSize so it automatically moves and resizes as the array expands or contracts. Optional line styling is applied before saving the file as ShapeLinkedToDynamicArray.xlsx.
class ShapeLinkDynamicArray
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the size of the dynamic array (5 rows x 3 columns)
            const int arrayRows = 5;
            const int arrayCols = 3;

            // Set a dynamic array formula in cell A1
            Cell arrayCell = sheet.Cells["A1"];
            arrayCell.Formula = $"=SEQUENCE({arrayRows},{arrayCols})";

            // Force calculation so the array result is materialized
            workbook.CalculateFormula();

            // Add a rectangle shape linked to the array range
            // Shapes.AddShape returns a Shape object in recent Aspose.Cells versions
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                arrayCell.Row,          // upper‑left row
                arrayCell.Column,       // upper‑left column
                0, 0,                   // offset within the cell
                arrayRows, arrayCols);  // height (rows) and width (columns)

            // Set the shape to move and size with cells
            shape.Placement = PlacementType.MoveAndSize;

            // Optional visual formatting (color properties may vary by version)
            shape.Line.Weight = 1.5;
            // Uncomment the following lines if the API version supports them
            // shape.Line.Color = Color.Blue;
            // shape.Fill.ForeColor = Color.FromArgb(128, 255, 255, 200);

            // Save the workbook
            string outputPath = "ShapeLinkedToDynamicArray.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
