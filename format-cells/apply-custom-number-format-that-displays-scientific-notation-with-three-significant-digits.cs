// Title: How to apply a custom scientific notation number format with three significant digits using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell format to scientific notation with three significant digits | C# example using Aspose.Cells to create a custom number format 0.00E+00 | How to display numbers in scientific notation with three sig figs using Aspose.Cells | Apply custom number format string 0.00E+00 to a cell in Aspose.Cells .NET | Saving workbook with scientific notation style using Aspose.Cells C#
// Tags: 0.00E+00 custom format Aspose.Cells | three significant digits scientific notation C# | cell style number formatting .NET | workbook save with custom style | Aspose.Cells number format API example

using Aspose.Cells;

// The example creates a new workbook, writes 12345.6789 to cell A1, defines a style with the custom number format "0.00E+00" to show three significant digits in scientific notation, applies the style to the cell, and saves the file as ScientificNotation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Put a sample number into cell A1
        var cell = sheet.Cells["A1"];
        cell.PutValue(12345.6789);

        // Create a style with custom scientific notation (three significant digits)
        var style = workbook.CreateStyle();
        style.Custom = "0.00E+00";

        // Apply the style to the cell
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("ScientificNotation.xlsx");
    }
}
