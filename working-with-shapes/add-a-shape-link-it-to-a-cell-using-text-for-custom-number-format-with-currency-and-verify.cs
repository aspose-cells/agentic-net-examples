// Title: Create a rectangle shape linked to a worksheet cell using the TEXT function for currency formatting and verify it with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# link shape text to a cell with TEXT formula for currency | How to display a custom formatted currency value inside a shape using Aspose.Cells | Verify shape text after calling workbook.CalculateFormula in Aspose.Cells .NET | Add a rectangle shape and bind its text to cell A1 with TEXT function in C# | Aspose.Cells example of shape formula reference with custom number format
// Tags: add rectangle shape Aspose.Cells | shape text formula reference Aspose.Cells | custom currency format TEXT function Aspose.Cells | calculate formulas retrieve shape text Aspose.Cells | save workbook linked shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a numeric value to cell A1, adds a rectangle shape, sets the shape's text to a TEXT formula that formats the cell value as currency, forces formula calculation, reads the displayed shape text to confirm it matches "$1,234.56", and saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234.56);

            // Add a rectangle shape to the worksheet
            // Parameters: drawing type, upper left row, upper left column, top, left, bottom, right
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,   // upper left row (zero‑based)
                2,   // upper left column (zero‑based)
                50,  // top (pixels)
                100, // left (pixels)
                150, // bottom (pixels)
                300  // right (pixels)
            );

            // Link the shape's text to cell A1 using the TEXT function for custom currency format
            shape.Text = "=TEXT(A1,\"$#,##0.00\")";

            // Force calculation of all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the displayed text of the shape after calculation
            string displayedText = shape.Text;

            // Verify that the shape shows the correctly formatted currency value
            Console.WriteLine("Shape text: " + displayedText);
            // Expected output: "$1,234.56"
            if (displayedText == "$1,234.56")
            {
                Console.WriteLine("Verification succeeded.");
            }
            else
            {
                Console.WriteLine("Verification failed.");
            }

            // Save the workbook (optional, for visual inspection)
            string outputPath = "ShapeLinkedToCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
