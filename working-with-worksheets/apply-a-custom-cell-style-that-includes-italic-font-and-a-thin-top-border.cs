// Title: Apply an italic font and a thin top border to a cell using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, defines a style with an italic font and a thin black top border, applies it to cell A1, and saves the file using Aspose.Cells. | Generate a method that receives a worksheet and a cell address, then sets the cell’s font to italic and adds a thin top border via Aspose.Cells styling. | Provide a step‑by‑step script to programmatically style a cell with italic text and a top border in a .NET application using Aspose.Cells.
// Common Searches: how to add italic text and top border to a specific cell in Aspose.Cells C# | Aspose.Cells C# example for setting thin top border with italic text | custom cell style with italic and border using Aspose.Cells for .NET | apply formatting to cell A1 in Aspose.Cells workbook C#
// Tags: italic text style Aspose.Cells C# | thin top edge border Aspose.Cells | custom cell style Aspose.Cells .NET | apply formatting to specific cell Aspose.Cells | set cell borders programmatically Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// Creates a new workbook, defines a style that makes the font italic and adds a thin black top border, applies the style to cell A1, writes "Sample", and saves the workbook as StyledCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Create a new style
        Style style = workbook.CreateStyle();

        // Set the font to italic
        style.Font.IsItalic = true;

        // Apply a thin top border
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.TopBorder].Color = Color.Black;

        // Write a value to cell A1 and apply the style
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Sample");
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("StyledCell.xlsx");
    }
}
