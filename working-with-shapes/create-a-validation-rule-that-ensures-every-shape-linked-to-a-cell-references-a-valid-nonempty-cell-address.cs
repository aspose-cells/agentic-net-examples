// Title: Aspose.Cells for .NET – Validate that every Shape LinkedCell points to a non‑empty, correctly formatted cell address
// Description: This example creates a workbook, adds rectangle shapes with valid, invalid and empty LinkedCell values, then scans all worksheets and shapes. It checks for null or whitespace, attempts to resolve the address via ws.Cells[linkedCell] to catch format errors, optionally verifies the cell is within the worksheet's used range, logs diagnostic messages, and saves the file without modification.
// Keywords: Aspose.Cells | .NET | C# | shape LinkedCell validation | cell address format check | empty LinkedCell detection | invalid address handling | worksheet range verification | Aspose.Cells example | GitHub sample
// Common Searches: how to validate shape LinkedCell in Aspose.Cells | detect empty or malformed LinkedCell address .NET | Aspose.Cells shape linked cell range check | C# code to verify shape cell reference | Aspose.Cells validate shape links to cells
// Developer Intent: Ensure every shape in a workbook has a non‑empty, syntactically valid LinkedCell reference and optionally lies within the used data range.
// Use Cases: Log shapes whose LinkedCell property is null, empty, or contains a malformed address. | Catch exceptions from ws.Cells[linkedCell] to identify invalid cell references. | Flag shapes that link to cells outside the worksheet's used rows or columns.
// AI Prompts: Generate a C# method that receives a Workbook and returns a list of shape names with empty, null, or invalid LinkedCell values using Aspose.Cells. | Provide code that validates each shape's LinkedCell against the worksheet's used range and throws a custom exception for out‑of‑range links. | Create a reusable routine that logs detailed messages for empty, malformed, or out‑of‑range LinkedCell values for all shapes across all worksheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkedCellValidation
{
    // This example creates a workbook, adds rectangle shapes with valid, invalid and empty LinkedCell values, then scans all worksheets and shapes. It checks for null or whitespace, attempts to resolve the address via ws.Cells[linkedCell] to catch format errors, optionally verifies the cell is within the worksheet's used range, logs diagnostic messages, and saves the file without modification.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape and link it to a valid cell (B2)
            Shape validShape = sheet.Shapes.AddRectangle(1, 1, 100, 50, 0, 0);
            validShape.LinkedCell = "$B$2";

            // Add another shape and intentionally set an invalid linked cell address
            Shape invalidShape = sheet.Shapes.AddRectangle(3, 3, 100, 50, 0, 0);
            invalidShape.LinkedCell = "InvalidAddress";

            // Add a shape with an empty linked cell
            Shape emptyShape = sheet.Shapes.AddRectangle(5, 5, 100, 50, 0, 0);
            emptyShape.LinkedCell = "";

            // Validate all shapes in all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Shape shape in ws.Shapes)
                {
                    string linkedCell = shape.LinkedCell;

                    // Check for null or whitespace
                    if (string.IsNullOrWhiteSpace(linkedCell))
                    {
                        Console.WriteLine($"Shape '{shape.Name}' has an empty LinkedCell reference.");
                        continue;
                    }

                    // Attempt to resolve the address; if it throws, the address is invalid
                    try
                    {
                        // The indexer creates the cell if it does not exist, but it validates the address format
                        Cell cell = ws.Cells[linkedCell];
                        // Optionally, you can check that the cell is within the worksheet's limits
                        if (cell.Row >= ws.Cells.MaxDataRow + 1 || cell.Column >= ws.Cells.MaxDataColumn + 1)
                        {
                            // This condition is rarely true because MaxDataRow/Column reflect used cells,
                            // but it demonstrates an additional range check if needed.
                            Console.WriteLine($"Shape '{shape.Name}' links to a cell outside the used range: {linkedCell}");
                        }
                        else
                        {
                            Console.WriteLine($"Shape '{shape.Name}' correctly links to cell {linkedCell}.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Shape '{shape.Name}' has an invalid LinkedCell reference '{linkedCell}'. Error: {ex.Message}");
                    }
                }
            }

            // Save the workbook (the validation does not modify the file)
            workbook.Save("ShapeLinkedCellValidationResult.xlsx", SaveFormat.Xlsx);
        }
    }
}
