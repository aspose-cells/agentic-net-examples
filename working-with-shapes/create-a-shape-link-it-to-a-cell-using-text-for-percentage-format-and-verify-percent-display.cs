// Title: Add a rectangle shape linked to a cell and display the cell's value as a formatted percentage using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a rectangle shape, sets its Text property to a TEXT formula referencing cell A1, and prints the evaluated shape text. | Show how to apply a built‑in percentage number format to a worksheet cell and bind a shape's text to that cell using the TEXT function in Aspose.Cells for .NET. | Demonstrate forcing formula calculation after linking a shape to a cell and then retrieving the shape's displayed text in a .NET workbook.
// Common Searches: aspnet aspose.cells link shape text to cell using TEXT function percentage format | c# how to display cell value as percent inside a shape with Aspose.Cells | aspose.cells shape text formula evaluation after workbook.CalculateFormula | add rectangle shape and bind it to a formatted cell value in Aspose.Cells for .NET | retrieve shape's displayed text after linking to a percentage cell in C#
// Tags: add rectangle shape aspose.cells | shape text linked to cell using TEXT function | apply percentage format to cell aspose.cells | force formula calculation workbook aspose.cells | read evaluated shape text c#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkExample
{
    // // Creates a workbook, puts a numeric value in A1, applies a built‑in percentage style, adds a rectangle shape, links the shape's Text to A1 with the TEXT function for percent formatting, forces formula calculation, reads the evaluated shape text, outputs it, and saves the file.
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

                // Set a numeric value in cell A1 (25.6%)
                Cell targetCell = sheet.Cells["A1"];
                targetCell.PutValue(0.256);

                // Apply percentage format to the cell (optional, for visual reference)
                Style percentStyle = workbook.CreateStyle();
                percentStyle.Number = 10; // Built‑in percentage format (0.00%)
                targetCell.SetStyle(percentStyle);

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, bottom, right
                // Here we place the shape roughly at rows 2‑5 and columns B‑D
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 1, 0, 0, 5, 3);

                // Link the shape's text to the cell using the TEXT function for percentage formatting
                // The formula will evaluate to a string like "25.60%"
                shape.Text = "=TEXT(A1,\"0.00%\")";

                // Force calculation of formulas so the shape's text is updated
                workbook.CalculateFormula();

                // Retrieve the evaluated text from the shape
                string shapeText = shape.Text;

                // Output the result to verify the percentage display
                Console.WriteLine("Shape linked text: " + shapeText); // Expected: "25.60%"

                // Save the workbook (optional, for inspection)
                workbook.Save("ShapeLinkedPercent.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
