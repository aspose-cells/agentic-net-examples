// Title: Add thin black borders to all sides of a cell with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example for setting thin black borders on a cell | How to apply borders on every side of an Excel cell using Aspose.Cells .NET | Aspose.Cells BorderType.Thin usage for cell border styling in C# | Set cell style borders to a black line in an Aspose.Cells workbook | Apply top, bottom, left, and right borders to a cell with the Aspose.Cells API
// Tags: Aspose.Cells thin black cell border | C# set cell borders Aspose.Cells | Excel cell border styling .NET | BorderType enumeration thin line Aspose.Cells | apply borders to Excel cell using Aspose.Cells | cell style border customization C#

using Aspose.Cells;
using System.Drawing;

// The sample creates a new Workbook, selects cell A1 on the first worksheet, retrieves its Style, sets a thin black line for the top, bottom, left, and right borders via the Borders collection, reapplies the Style to the cell, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Target cell (A1) to apply borders
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Sample"); // optional content

        // Retrieve the cell's style
        Style style = cell.GetStyle();

        // Define a thin black line for each border side
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.TopBorder].Color = Color.Black;

        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.BottomBorder].Color = Color.Black;

        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.LeftBorder].Color = Color.Black;

        style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.RightBorder].Color = Color.Black;

        // Apply the modified style back to the cell
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
