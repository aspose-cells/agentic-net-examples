// Title: C# – Link a Shape to a Cell with CONCAT and Display the Result using Aspose.Cells
// Description: Creates a workbook, fills A1 and B1, sets C1 to a CONCAT formula, calculates it, adds a rectangle shape, links the shape to C1, refreshes the shape text, verifies the displayed value, and saves the file.
// Keywords: Aspose.Cells | C# shape linking | LinkedCell property | CONCAT formula | rectangle shape text | Excel shape automation
// Common Searches: Aspose.Cells link shape to cell C# | display CONCAT result in Excel shape | update shape text after formula calculation | how to bind a rectangle to a cell with Aspose.Cells | verify shape Text property matches linked cell
// Developer Intent: Show how to bind a rectangle shape to a cell that contains a CONCAT formula and ensure the shape shows the concatenated string.
// Use Cases: Dynamic dashboards where shape captions reflect combined cell values. | Automated report generation with shape labels driven by text formulas. | Maintaining visual consistency in exported workbooks by syncing shapes with formula results.
// AI Prompts: Provide C# code that links a rectangle shape to a cell containing a CONCAT formula and refreshes its displayed text with Aspose.Cells. | Demonstrate how to verify that a shape's Text property equals the value of its linked cell after formula evaluation. | Explain how to apply the LinkedCell property to multiple shapes that use different text functions in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, fills A1 and B1, sets C1 to a CONCAT formula, calculates it, adds a rectangle shape, links the shape to C1, refreshes the shape text, verifies the displayed value, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill cells A1 and B1 with sample text
            worksheet.Cells["A1"].PutValue("Hello");
            worksheet.Cells["B1"].PutValue("World");

            // Set C1 to a CONCAT formula that joins A1 and B1 with a space
            worksheet.Cells["C1"].Formula = "CONCAT(A1,\" \",B1)";

            // Calculate the formula so C1 contains the concatenated result
            workbook.CalculateFormula();

            // Add a rectangle shape to the sheet
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 200, 50, 0, 0);

            // Link the shape to cell C1 (the cell that holds the concatenated string)
            shape.LinkedCell = "C1";

            // Refresh the shape's displayed value based on the linked cell
            shape.UpdateSelectedValue();

            // Verify: read the value from the linked cell
            string linkedCellValue = worksheet.Cells["C1"].StringValue;
            Console.WriteLine("Linked cell value: " + linkedCellValue);

            // Verify: read the text displayed by the shape using the Text property
            string shapeText = shape.Text ?? string.Empty;
            Console.WriteLine("Shape displayed text: " + shapeText);

            // Save the workbook
            workbook.Save("ShapeLinkedCellConcat.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
