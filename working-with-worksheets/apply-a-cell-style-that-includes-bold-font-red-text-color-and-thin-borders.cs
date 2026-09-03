// Title: Apply a bold red font and thin borders to a single cell with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style with a bold red font and thin borders on all sides, then applies it to cell A1 using Aspose.Cells. | Show how to define a custom cell style in Aspose.Cells, set Font.IsBold, Font.Color, and Borders.LineStyle, and assign it to a specific cell. | Provide a complete example that formats a cell with bold red text and thin borders, writes a value, and saves the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# how to set bold red font and thin borders for a single cell | C# example applying custom style to cell A1 with Aspose.Cells | formatting cell with font color and border style using Aspose.Cells .NET | apply thin border and red bold text to a cell in an Excel workbook with Aspose.Cells
// Tags: bold red font cell style Aspose.Cells C# | thin border formatting Aspose.Cells | custom cell style application .NET workbook | set cell font color and border Aspose.Cells | apply style to specific cell Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// The sample creates a new workbook, defines a style that makes the font bold and red and adds thin borders on all sides, applies this style to cell A1, writes "Styled Text", and saves the file as StyledWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Create a new style
        Style style = workbook.CreateStyle();

        // Set font to bold and red color
        style.Font.IsBold = true;
        style.Font.Color = Color.Red;

        // Set thin borders on all sides
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

        // Apply the style to cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Styled Text");
        cell.SetStyle(style);

        // Save the workbook (lifecycle rule)
        workbook.Save("StyledWorkbook.xlsx");
    }
}
