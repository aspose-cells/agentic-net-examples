// Title: Create a rectangle shape in Aspose.Cells for .NET and bind its text to a cell using the TEXT function with a custom currency format
// AI Prompts: Add a rectangle shape to a worksheet and assign its Text property a TEXT formula that references cell A1 with the "$#,##0.00" format. | Run workbook.CalculateFormula() and then read the shape's Text property to obtain the evaluated currency string. | Persist the workbook as an XLSX file after the shape text has been linked to the formatted cell value.
// Common Searches: Aspose.Cells display a cell's formatted currency value inside a shape using TEXT formula | link rectangle shape text to a cell with custom currency format in C# | evaluate shape formulas after CalculateFormula in Aspose.Cells .NET | set custom number format for a cell and reference it in shape text Aspose.Cells example
// Tags: rectangle shape TEXT formula Aspose.Cells | custom currency number format cell .NET | shape text evaluation CalculateFormula | binding shape text to cell value Aspose.Cells | saving workbook with linked shape XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a numeric value to cell A1, applies a custom currency number format, adds a rectangle shape, links the shape's text to the cell using the TEXT function, recalculates formulas to evaluate the shape's displayed text, prints the result, and saves the workbook as an XLSX file.
class ShapeLinkExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a numeric value in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234.56);

            // Apply a custom number format with currency to the cell
            Style style = cell.GetStyle();
            style.Custom = "$#,##0.00";
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top, left, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 1, 5, 5, 30, 150);

            // Link the shape's text to the cell using the TEXT function with the same format
            shape.Text = "=TEXT(A1,\"$#,##0.00\")";

            // Recalculate formulas so the shape's text is evaluated
            workbook.CalculateFormula();

            // Verify the shape's displayed text
            string displayedText = shape.Text;
            Console.WriteLine("Shape text after evaluation: " + displayedText);
            // Expected output: $1,234.56

            // Save the workbook (optional, just to demonstrate lifecycle)
            workbook.Save("ShapeLinkExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
