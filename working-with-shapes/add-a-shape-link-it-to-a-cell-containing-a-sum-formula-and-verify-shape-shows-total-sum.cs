// Title: C# – Add a Rectangle Shape Linked to a SUM Formula Cell Using Aspose.Cells
// Description: Demonstrates how to create a workbook, fill cells A1‑A5, set a SUM formula in A6, calculate the sheet, add a rectangle shape, link the shape to cell A6, refresh the shape’s displayed value, and save the file. The example verifies that the shape shows the same total as the formula.
// Keywords: Aspose.Cells C# shape linking | link shape to cell formula | rectangle shape SUM result | update shape value Aspose.Cells | Excel shape linked cell .NET | display formula result in shape
// Common Searches: Aspose.Cells link shape to formula cell C# | how to display SUM result in a worksheet shape | update shape value after linking to a cell Aspose.Cells | C# add rectangle shape linked to cell | verify shape shows calculated total Aspose.Cells
// Developer Intent: Link a worksheet shape to a cell that contains a SUM formula and confirm the shape displays the calculated total.
// Use Cases: Financial dashboards where shapes automatically reflect totals calculated by formulas. | Automated Excel reports that use linked shapes to highlight key metrics without manual edits. | Dynamic workbooks where shapes act as visual indicators for calculation results.
// AI Prompts: Show how to change a linked shape’s font size and color after linking it to a cell with Aspose.Cells. | Provide code to link multiple shapes to different formula cells and refresh all displayed values in one workbook. | Explain error handling when the linked cell contains an invalid or #REF! formula while updating the shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, fill cells A1‑A5, set a SUM formula in A6, calculate the sheet, add a rectangle shape, link the shape to cell A6, refresh the shape’s displayed value, and save the file. The example verifies that the shape shows the same total as the formula.
class ShapeLinkedCellDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells A1:A5 with sample numbers
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].PutValue(30);
        worksheet.Cells["A4"].PutValue(40);
        worksheet.Cells["A5"].PutValue(50);

        // Set a SUM formula in cell A6 that sums A1:A5
        worksheet.Cells["A6"].Formula = "=SUM(A1:A5)";

        // Calculate the workbook to evaluate the formula
        workbook.CalculateFormula();

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top, left, height, width
        Shape rectangle = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 30);

        // Link the shape to the cell containing the SUM formula (A6)
        rectangle.LinkedCell = "$A$6";

        // Update the shape's displayed value based on the linked cell
        rectangle.UpdateSelectedValue();

        // Retrieve the calculated sum from the linked cell
        double sumValue = worksheet.Cells["A6"].DoubleValue;

        // Output verification information
        Console.WriteLine("Linked cell for shape: " + rectangle.LinkedCell);
        Console.WriteLine("Calculated SUM in A6: " + sumValue);
        Console.WriteLine("Shape should display the same total sum.");

        // Save the workbook to a file
        workbook.Save("ShapeLinkedCellDemo.xlsx");
    }
}
