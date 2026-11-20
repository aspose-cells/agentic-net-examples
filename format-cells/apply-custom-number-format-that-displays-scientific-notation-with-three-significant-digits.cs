// Title: Apply a custom scientific notation number format with three significant digits in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set custom number format 0.00E+00 for scientific notation | How to display three significant digits in Excel using Aspose.Cells style | C# code to apply scientific notation format to a worksheet cell with Aspose | Create a custom number format for scientific notation in a .NET workbook | Aspose.Cells format cell as scientific notation with three sig figs
// Tags: Aspose.Cells scientific notation style C# | three significant digits formatting .NET | 0.00E+00 pattern for Excel cells | apply cell style programmatically Aspose | Excel number format with sig figs using Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, writes the value 12345.6789 to cell A1, defines a custom number format "0.00E+00" to show three significant digits in scientific notation, applies this style to the cell, and saves the file as ScientificNotation.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        // Insert a numeric value into cell A1
        var cell = sheet.Cells["A1"];
        cell.PutValue(12345.6789);

        // Define a custom number format for scientific notation with three significant digits
        var style = workbook.CreateStyle();
        style.Custom = "0.00E+00";

        // Apply the custom style to the cell
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("ScientificNotation.xlsx");
    }
}
