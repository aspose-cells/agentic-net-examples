// Title: Apply an italic font and a thin top border to a specific cell with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style with an italic typeface and a thin black top border, applies it to cell A1, and saves the workbook using Aspose.Cells. | Demonstrate how to define a custom cell style containing slanted text and a top edge border in Aspose.Cells for .NET and assign it to a chosen cell.
// Common Searches: asp.net aspose.cells set slanted text and top edge border for a single cell | c# example adding thin top border to Excel cell with Aspose.Cells | how to create custom style with italic font in Aspose.Cells workbook | apply border only to top side of a cell using Aspose.Cells .NET
// Tags: italic font style Aspose.Cells | thin top border formatting C# | custom cell style creation Aspose.Cells .NET | set cell top edge border programmatically | Excel cell formatting with font and border Aspose

using Aspose.Cells;
using System.Drawing;

// The program creates a new workbook, writes "Sample Text" to cell A1, defines a style with italic text and a thin black top border, applies this style to the cell, and saves the file as StyledCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the cell to style (e.g., A1) and put a sample value
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Sample Text");

        // Create a new style object
        Style style = workbook.CreateStyle();

        // Set the font to italic
        style.Font.IsItalic = true;

        // Configure a thin top border
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.TopBorder].Color = Color.Black;

        // Apply the style to the cell
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("StyledCell.xlsx");
    }
}
