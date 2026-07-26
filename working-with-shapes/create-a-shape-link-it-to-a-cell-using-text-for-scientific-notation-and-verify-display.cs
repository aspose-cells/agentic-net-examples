// Title: C# – Link a Rectangle Shape to a TEXT‑formatted Scientific Notation Cell with Aspose.Cells
// Description: Shows how to put a number in A1, format it as scientific notation using the TEXT function in B1, attach a rectangle shape to B1, refresh the shape text, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | shape linked cell | rectangle shape | TEXT function | scientific notation | SetLinkedCell | UpdateSelectedValue | formula calculation
// Common Searches: Aspose.Cells link shape to cell | C# shape displays TEXT formula result | display scientific notation in shape Aspose.Cells | SetLinkedCell rectangle shape .NET | refresh linked shape after formula
// Developer Intent: Create a rectangle shape, bind it to a cell that returns a TEXT‑formatted scientific notation string, and confirm the shape displays that text.
// Use Cases: Dynamic labels that automatically reflect formatted numeric results. | Automated report generation where shapes serve as data‑driven captions. | Testing that linked shapes correctly inherit formula output before saving.
// AI Prompts: Generate C# code with Aspose.Cells to add a rectangle shape, link it to a cell using the TEXT function for scientific notation, update the shape, and print its Text property. | Explain the effect of each parameter in SetLinkedCell when linking a shape to a worksheet cell in Aspose.Cells. | Outline the steps to recalculate formulas and refresh a linked shape's displayed value in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsShapeLinkedCellDemo
{
    // Shows how to put a number in A1, format it as scientific notation using the TEXT function in B1, attach a rectangle shape to B1, refresh the shape text, and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a numeric value in cell A1
            worksheet.Cells["A1"].PutValue(123456789);

            // In cell B1 use the TEXT function to format A1 in scientific notation
            // The formula will produce a string like "1.23E+08"
            worksheet.Cells["B1"].Formula = "TEXT(A1,\"0.00E+00\")";

            // Recalculate formulas so B1 contains the formatted text
            workbook.CalculateFormula();

            // Add a rectangle shape that will act as a text box
            // Parameters: upper left row, upper left column, top, left, height, width
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 30);

            // Link the shape to cell B1 (which holds the scientific notation text)
            // isR1C1 = false (A1 style), isLocal = true (locale aware)
            shape.SetLinkedCell("$B$1", false, true);

            // Update the shape's displayed value from the linked cell
            shape.UpdateSelectedValue();

            // Verify the displayed text by reading the Shape.Text property
            Console.WriteLine("Shape text (should be scientific notation): " + shape.Text);

            // Save the workbook to a file
            workbook.Save("ShapeLinkedCellScientificNotation.xlsx");
        }
    }
}
