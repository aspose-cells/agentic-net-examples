// Title: Apply a bold red font and thin borders to a single Excel cell with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a custom Style in Aspose.Cells with a bold red font and thin borders on all sides, then applies it to cell A1. | Show how to use StyleFlag in Aspose.Cells to apply all style attributes (font, color, borders) to a specific cell. | Provide an example that saves the workbook after styling the cell as an .xlsx file using Aspose.Cells.
// Common Searches: Aspose.Cells C# set cell font to bold red and add thin borders | How to apply a style with borders to a single cell using Aspose.Cells .NET | Using StyleFlag to apply full formatting to a cell in Aspose.Cells C# | Save styled Excel workbook after applying custom cell style with Aspose.Cells
// Tags: custom cell style bold red font Aspose.Cells | thin border formatting Aspose.Cells C# | StyleFlag apply all attributes .NET | save styled workbook as xlsx Aspose.Cells | apply formatting to single Excel cell C#

using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, defines a style with bold red text and thin borders, applies it to cell A1 using a StyleFlag that enables all attributes, and saves the file as StyledCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get cell A1 and set a sample value
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Styled Text");

        // Create a new style
        Style style = workbook.CreateStyle();

        // Set font to bold and red color
        style.Font.IsBold = true;
        style.Font.Color = Color.Red;

        // Apply thin borders on all sides
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

        // Apply the style to the cell
        StyleFlag flag = new StyleFlag();
        flag.All = true; // apply all style attributes
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("StyledCell.xlsx");
    }
}
