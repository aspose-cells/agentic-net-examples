// Title: Apply a bold red font and thin borders to an Excel cell with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set cell font to bold red and add thin borders | programmatically style Excel cell with bold red text using Aspose.Cells | C# example for applying thin borders to a cell in Aspose.Cells workbook | how to create and assign a custom style to a worksheet cell in Aspose.Cells .NET | set cell style with red font and borders in Aspose.Cells for .NET
// Tags: Aspose.Cells bold red font style | apply thin borders using CellBorderType.Thin | C# workbook.CreateStyle custom cell style | SetStyle method cell formatting Aspose.Cells | Excel cell formatting with Aspose.Cells .NET

using System;
using Aspose.Cells;
using System.Drawing;

// The example creates a new workbook, writes "Styled Text" to cell A1, defines a style with a bold red font and thin borders on all sides, applies this style to the cell, and saves the workbook as StyledWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Access a specific cell (e.g., A1)
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Styled Text");

        // Create a new style object
        Style style = workbook.CreateStyle();

        // Set font to bold
        style.Font.IsBold = true;

        // Set font color to red
        style.Font.Color = Color.Red;

        // Apply thin borders on all sides
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

        // Apply the style to the cell
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("StyledWorkbook.xlsx");
    }
}
