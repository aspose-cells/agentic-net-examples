// Title: How to add a thin black border on all sides of a cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style with a thin black border on the top, bottom, left, and right and applies it to a specified cell in an Aspose.Cells workbook. | Provide a reusable C# method named ApplyThinBlackBorder that takes a Worksheet and a cell address, sets a uniform thin black border using Aspose.Cells, and returns the styled cell. | Show the steps to save the workbook as an .xlsx file after applying the border style to a cell with Aspose.Cells.
// Common Searches: Aspose.Cells C# set uniform thin black border for a single cell | How to apply borders on all sides of a cell with Aspose.Cells .NET | C# example adding thin black borders to cell A1 using Aspose.Cells | Aspose.Cells style object border color and line style tutorial | Saving workbook after formatting cell borders in Aspose.Cells C#
// Tags: apply thin black border Aspose.Cells C# | cell style borders Aspose.Cells .NET | uniform cell border formatting Aspose.Cells | save workbook after styling cell Aspose.Cells | CellBorderType.Thin usage example

using Aspose.Cells;
using System.Drawing;

// // Demonstrates creating a workbook, defining a style with thin black borders on all four sides, applying it to cell A1, and saving the file as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the cell to apply borders (e.g., A1)
        Cell cell = sheet.Cells["A1"];

        // Create a style object
        Style style = workbook.CreateStyle();

        // Set thin black border on all sides
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.TopBorder].Color = Color.Black;

        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.BottomBorder].Color = Color.Black;

        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.LeftBorder].Color = Color.Black;

        style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.RightBorder].Color = Color.Black;

        // Apply the style to the cell
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
