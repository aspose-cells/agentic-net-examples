// Title: Validate Shape LinkedCell Addresses with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds rectangle shapes with linked cells (valid, empty, malformed), iterates through each shape, verifies that the LinkedCell string is non‑empty and resolves to a real cell, clears invalid references, and saves the file.
// Keywords: Aspose.Cells shape linked cell validation | C# linked cell address check | clear invalid shape links | non‑empty LinkedCell property | Aspose.Cells workbook cleanup
// Common Searches: how to verify shape linked cell address Aspose.Cells C# | remove invalid linked cells from shapes Aspose.Cells | validate and clear shape LinkedCell in .NET | detect empty or malformed linked cell in Aspose.Cells | shape LinkedCell validation example
// Developer Intent: Ensure every shape’s LinkedCell points to a non‑empty, syntactically correct cell and automatically clear any invalid links.
// Use Cases: Pre‑save validation of all shapes in a worksheet to prevent runtime errors. | Cleaning up templates that contain shapes with missing or corrupt linked cells. | Automated post‑processing of generated reports to remove bad shape references.
// AI Prompts: Generate a C# method that returns a list of shape names with invalid LinkedCell values in a given worksheet. | Write code to log and optionally fix shapes that have empty or malformed LinkedCell strings using Aspose.Cells. | Create an NUnit test suite that confirms the validator correctly identifies valid, empty, and malformed linked cell addresses.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// Creates a workbook, adds rectangle shapes with linked cells (valid, empty, malformed), iterates through each shape, verifies that the LinkedCell string is non‑empty and resolves to a real cell, clears invalid references, and saves the file.
class ShapeLinkedCellValidator
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add shapes with various linked cell values
            Shape shapeValid = worksheet.Shapes.AddRectangle(1, 1, 100, 100, 0, 0);
            shapeValid.SetLinkedCell("$A$1", false, false); // valid address

            Shape shapeEmpty = worksheet.Shapes.AddRectangle(2, 2, 100, 100, 0, 0);
            shapeEmpty.SetLinkedCell(string.Empty, false, false); // empty address

            Shape shapeInvalid = worksheet.Shapes.AddRectangle(3, 3, 100, 100, 0, 0);
            shapeInvalid.SetLinkedCell("Invalid!Ref", false, false); // malformed address

            // Validate each shape's linked cell
            foreach (Shape shape in worksheet.Shapes)
            {
                string linkedCell = shape.LinkedCell;
                bool isValid = false;

                // Non‑empty check
                if (!string.IsNullOrEmpty(linkedCell))
                {
                    try
                    {
                        // Attempt to retrieve the cell; if no exception, address is syntactically valid
                        Cell cell = worksheet.Cells[linkedCell];
                        // Additional check: ensure the cell object exists
                        isValid = cell != null;
                    }
                    catch
                    {
                        // Any exception means the address is invalid
                        isValid = false;
                    }
                }

                if (!isValid)
                {
                    Console.WriteLine($"Shape \"{shape.Name}\" has invalid linked cell '{linkedCell}'. Clearing link.");
                    shape.LinkedCell = string.Empty; // Clear invalid reference
                }
                else
                {
                    Console.WriteLine($"Shape \"{shape.Name}\" linked to valid cell '{linkedCell}'.");
                }
            }

            // Save the workbook with validated shapes
            workbook.Save("ShapeLinkedCellValidated.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
