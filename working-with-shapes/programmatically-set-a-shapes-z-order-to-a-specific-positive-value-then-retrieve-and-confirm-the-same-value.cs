// Title: Aspose.Cells for .NET – Set and Verify Shape ZOrderPosition (Z‑Order)
// Description: Demonstrates how to add a rectangle shape to a worksheet, assign a positive ZOrderPosition value, read it back, confirm the assignment, and save the workbook as ZOrderDemo.xlsx using C# and Aspose.Cells.
// Keywords: Aspose.Cells C# shape ZOrderPosition | set shape Z‑order Aspose.Cells | retrieve ZOrderPosition property | worksheet shape stacking order | .NET spreadsheet shape order | Aspose.Cells shape Z‑order example
// Common Searches: how to set shape ZOrderPosition in Aspose.Cells | C# retrieve shape Z‑order after setting it | Aspose.Cells change shape stacking order | verify shape ZOrderPosition value programmatically
// Developer Intent: Assign a specific Z‑order to a worksheet shape and ensure the value is stored correctly.
// Use Cases: Control visual layering of multiple shapes by setting explicit ZOrderPosition values. | Validate shape order before exporting a workbook to maintain design fidelity. | Adjust shape stacking dynamically based on user interaction or data conditions.
// AI Prompts: Write C# code that adds three different shapes to a worksheet and arranges them using ZOrderPosition so that shape A appears on top, B in the middle, and C at the bottom. | Explain the relationship between ZOrderPosition, BringToFront, and SendToBack methods in Aspose.Cells. | Show how to handle exceptions when an invalid (negative or out‑of‑range) ZOrderPosition is assigned to a shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderDemo
{
    // Demonstrates how to add a rectangle shape to a worksheet, assign a positive ZOrderPosition value, read it back, confirm the assignment, and save the workbook as ZOrderDemo.xlsx using C# and Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddRectangle(5, 5, 0, 0, 100, 100);

            // Set the Z-order position to a specific positive value (e.g., 5)
            shape.ZOrderPosition = 5;

            // Retrieve the Z-order position
            int retrievedZOrder = shape.ZOrderPosition;

            // Confirm that the set and retrieved values are the same
            if (retrievedZOrder == 5)
            {
                Console.WriteLine("Z-order successfully set and verified: " + retrievedZOrder);
            }
            else
            {
                Console.WriteLine("Z-order verification failed. Expected 5, got " + retrievedZOrder);
            }

            // Save the workbook to a file
            workbook.Save("ZOrderDemo.xlsx");
        }
    }
}
