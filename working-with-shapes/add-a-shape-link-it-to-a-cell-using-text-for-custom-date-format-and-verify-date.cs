// Title: Add a rectangle shape linked to a cell with a TEXT formula for custom date formatting and verify the displayed date using Aspose.Cells for .NET
// AI Prompts: Insert a rectangle shape on the first worksheet, assign its Text property the formula =TEXT(A1,"yyyy-MM-dd"), invoke workbook.CalculateFormula(), then read shape.Text and compare it to the formatted value of cell A1 to confirm the dates match. | Change the linked cell to B2, write a new DateTime value, update the shape's Text formula to =TEXT(B2,"dd MMM yyyy"), recalculate formulas, and verify that shape.Text reflects the new formatted date.
// Common Searches: Aspose.Cells C# shape text linked to cell using TEXT function | how to bind shape text to a worksheet cell for date formatting | checking that shape displays same date as cell after calculation | example of shape with formula-driven text in Aspose.Cells | using TEXT function inside shape to format dates in .NET
// Tags: shape insertion Aspose.Cells | text formula linking for shapes | date formatting via TEXT formula in shape | formula evaluation for shape content | date verification between shape and cell | shape-to-cell data binding Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a workbook, writes a DateTime to cell A1, adds a rectangle shape, links the shape's text to A1 using the TEXT function with a custom "yyyy-MM-dd" format, forces formula calculation, retrieves the evaluated text, compares it to the cell's formatted value, records the verification result in B1, and saves the file as LinkedShapeWithDate.xlsx.
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

                // Put a date value into cell A1
                sheet.Cells["A1"].PutValue(new DateTime(2023, 12, 25));

                // Add a rectangle shape to the worksheet
                Shape shape = sheet.Shapes.AddShape(
                    MsoDrawingType.Rectangle, // shape type
                    1,    // upper left row
                    0,    // upper left column
                    0,    // upper left row offset in pixels
                    0,    // upper left column offset in pixels
                    100,  // shape width in pixels
                    30    // shape height in pixels
                );

                // Link the shape's text to cell A1 using the TEXT function for a custom date format
                shape.Text = "=TEXT(A1,\"yyyy-MM-dd\")";

                // Force calculation of all formulas so the TEXT function is evaluated
                workbook.CalculateFormula();

                // Retrieve the evaluated text from the shape
                string shapeText = shape.Text;

                // Verify that the shape displays the expected formatted date
                string expectedDate = sheet.Cells["A1"].DateTimeValue.ToString("yyyy-MM-dd");
                bool isDateCorrect = shapeText == expectedDate;

                // Write verification result back to the workbook for reference
                sheet.Cells["B1"].PutValue(isDateCorrect ? "Date verified" : "Date mismatch");

                // Save the workbook
                string outputPath = "LinkedShapeWithDate.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
