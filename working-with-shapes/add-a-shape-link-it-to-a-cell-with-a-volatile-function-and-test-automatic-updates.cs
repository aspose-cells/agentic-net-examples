// Title: Add a rectangle shape linked to a volatile NOW() cell and automatically refresh its text with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example linking a shape to a cell with a volatile NOW() formula | How to automatically update shape text after workbook recalculation in Aspose.Cells .NET | Using AlternativeText property to bind a rectangle shape to a cell address in Aspose.Cells | Refresh shape content when volatile functions like NOW() change in C# | Step-by-step guide to sync shape text with cell value in an Aspose.Cells workbook
// Tags: link shape to cell address Aspose.Cells C# | rectangle shape alternativetext binding | auto refresh shape text after volatile formula | recalculate workbook formulas update shape | NOW() function synchronization with shape | Aspose.Cells shape text synchronization

using System;
using System.Drawing;
using System.IO;
using System.Threading;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkExample
{
    // The C# program creates a workbook, writes a volatile NOW() formula to cell B2, adds a rectangle shape, stores the cell address in the shape's AlternativeText, sets the shape's text to the current cell value, saves the file, pauses, recalculates formulas, updates the shape text with the new value, and saves the updated workbook, demonstrating automatic shape text refresh linked to a volatile function.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Put a volatile function (NOW) into cell B2
                Cell volatileCell = sheet.Cells["B2"];
                volatileCell.Formula = "=NOW()";

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, bottom, right
                int upperLeftRow = 5;
                int upperLeftColumn = 2;
                int lowerRightRow = 7;
                int lowerRightColumn = 4;
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle,
                    upperLeftRow,
                    upperLeftColumn,
                    0,
                    0,
                    lowerRightRow,
                    lowerRightColumn);

                // Set a visible fill color for the shape (optional)
                shape.FillFormat.ForeColor = Color.LightBlue;

                // Link the shape to the volatile cell by storing the cell address in AlternativeText
                shape.AlternativeText = "B2";

                // Initial calculation to evaluate the volatile function
                workbook.CalculateFormula();

                // Retrieve the current value of the volatile cell
                string cellValue = volatileCell.StringValue;

                // Set the shape's text to the cell's value
                shape.Text = $"Current Time: {cellValue}";

                // Save the workbook after initial update
                string initialPath = "ShapeLinkedToVolatileCell_Initial.xlsx";
                workbook.Save(initialPath);
                Console.WriteLine($"Saved initial workbook to '{Path.GetFullPath(initialPath)}'.");

                // Simulate a later time when the workbook is recalculated (e.g., after some minutes)
                Thread.Sleep(2000); // pause 2 seconds
                workbook.CalculateFormula();

                // Update the shape's text again to reflect the new volatile value
                cellValue = volatileCell.StringValue;
                shape.Text = $"Updated Time: {cellValue}";

                // Save the workbook after the automatic update
                string updatedPath = "ShapeLinkedToVolatileCell_Updated.xlsx";
                workbook.Save(updatedPath);
                Console.WriteLine($"Saved updated workbook to '{Path.GetFullPath(updatedPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
