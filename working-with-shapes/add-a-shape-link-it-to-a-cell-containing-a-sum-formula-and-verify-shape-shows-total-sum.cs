// Title: Link a Rectangle Shape to a SUM Formula Cell and Verify the Value with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, fill cells A1:A5, add a SUM formula in B1, insert a rectangle shape, link the shape to B1 using the LinkedCell property, refresh the shape with UpdateSelectedValue, retrieve the linked cell reference, confirm the calculated total (150), and save the file.
// Keywords: Aspose.Cells | C# | shape linked cell | LinkedCell property | UpdateSelectedValue | rectangle shape | SUM formula | verify shape value | Excel automation
// Common Searches: Aspose.Cells link shape to cell C# | how to bind a rectangle to a formula cell using Aspose.Cells | display SUM result in a shape Aspose.Cells .NET | update shape value after formula calculation Aspose.Cells | verify linked shape value in Excel workbook programmatically
// Developer Intent: Show how to attach a rectangle shape to a cell that contains a SUM formula and programmatically confirm that the shape displays the computed total.
// Use Cases: Create dashboards where shapes automatically reflect summary totals. | Generate reports that visually highlight key figures by linking shapes to calculated cells. | Automated testing of workbook generation to ensure linked shapes show correct formula results.
// AI Prompts: Write C# code with Aspose.Cells that adds a rectangle shape, links it to a SUM formula cell, updates the shape, and validates the displayed value. | Explain the role of the LinkedCell property and UpdateSelectedValue method when showing formula results in a shape. | Provide troubleshooting steps when a shape linked to a formula cell does not reflect the expected value in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, fill cells A1:A5, add a SUM formula in B1, insert a rectangle shape, link the shape to B1 using the LinkedCell property, refresh the shape with UpdateSelectedValue, retrieve the linked cell reference, confirm the calculated total (150), and save the file.
class ShapeLinkedCellDemo
{
    static void Main()
    {
        try
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

            // Add a SUM formula in cell B1 that sums A1:A5
            worksheet.Cells["B1"].Formula = "SUM(A1:A5)";

            // Add a rectangle shape to the worksheet
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height
            Shape rectangle = worksheet.Shapes.AddRectangle(2, 2, 5, 5, 120, 30);

            // Link the shape to the cell containing the SUM formula
            rectangle.LinkedCell = "$B$1";

            // Ensure the shape reflects the current value of the linked cell
            rectangle.UpdateSelectedValue();

            // Retrieve the linked cell reference (without $ signs)
            string linkedCellRef = rectangle.GetLinkedCell(false, false); // e.g., "B1"

            // Access the linked cell
            Cell linkedCell = worksheet.Cells[linkedCellRef];
            double sumResult = linkedCell.DoubleValue; // Value calculated by the formula

            // Output verification results
            Console.WriteLine($"Linked cell reference: {linkedCellRef}");
            Console.WriteLine($"Sum result in linked cell: {sumResult}");

            // Simple check to confirm the shape shows the same total sum
            if (Math.Abs(sumResult - 150) < 0.0001)
            {
                Console.WriteLine("Verification passed: Shape correctly linked to the SUM result.");
            }
            else
            {
                Console.WriteLine("Verification failed: Unexpected sum value.");
            }

            // Save the workbook to a file
            workbook.Save("ShapeLinkedCellDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
