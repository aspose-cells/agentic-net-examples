// Title: How to link a rectangle shape to a cell formatted with the TEXT function in Aspose.Cells for .NET
// AI Prompts: Create a rectangle shape on a worksheet and assign its Text property to the string value of a cell that contains a TEXT formula, then save the workbook. | Calculate worksheet formulas, retrieve the formatted result from a cell, and bind that result to a shape’s displayed text with custom alignment and font styling using Aspose.Cells.
// Common Searches: Aspose.Cells set shape text to result of TEXT formula in another cell | C# link rectangle shape to formatted number using TEXT function Aspose.Cells | How to display a cell's TEXT function output inside a shape with Aspose.Cells .NET
// Tags: Aspose.Cells shape text binding to cell value | rectangle shape linked to TEXT function output | set shape text from formatted cell Aspose.Cells | C# calculate formulas and bind result to shape | Aspose.Cells formatted number display in shape

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes a raw number to A1, formats it as currency in B1 using the TEXT function, calculates formulas, adds a rectangle shape, sets the shape's text to the formatted string from B1, applies center alignment and bold styling, and saves the file as LinkedShapeWithFormattedNumber.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a raw number in cell A1
            Cell rawNumberCell = sheet.Cells["A1"];
            rawNumberCell.PutValue(12345.6789);

            // In cell B1, use the TEXT function to format the number as currency
            Cell formattedCell = sheet.Cells["B1"];
            formattedCell.Formula = @"=TEXT(A1,""$#,##0.00"")";

            // Calculate formulas so that B1 contains the formatted string
            workbook.CalculateFormula();

            // Add a rectangle shape to the worksheet
            // Parameters: shape type, upper left row, upper left column, row offset, column offset, height, width
            Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 2, 2, 0, 0, 60, 200);

            // Link the shape's displayed text to the formatted value in B1
            shape.Text = formattedCell.StringValue;

            // Optionally, set the shape's text alignment and style
            shape.TextHorizontalAlignment = TextAlignmentType.Center;
            shape.TextVerticalAlignment = TextAlignmentType.Center;
            shape.Font.Size = 12;
            shape.Font.IsBold = true;

            // Save the workbook
            workbook.Save("LinkedShapeWithFormattedNumber.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
