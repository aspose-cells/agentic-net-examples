// Title: Apply Italic Font and Thin Top Border with Aspose.Cells for .NET (C#)
// Description: Shows how to build a custom Style in Aspose.Cells, set the font to italic, add a thin black top border, use a StyleFlag to limit the changes to those two attributes, apply the style to cell A1, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# | custom cell style | italic font | top border | thin border | StyleFlag | Excel formatting | Workbook.Save | XLSX output
// Common Searches: Aspose.Cells set italic font C# | Aspose.Cells thin top border example | How to use StyleFlag in Aspose.Cells | Apply custom style to a single cell Aspose.Cells | C# Excel cell formatting with Aspose
// Developer Intent: The developer wants to format a specific cell so that its text is italic and the cell has a thin black top border, without altering other style properties.
// Use Cases: Create header cells that stand out with italic text and a separating top line. | Highlight subtotal rows by adding a thin top border and italic font for visual emphasis. | Define a reusable style for table headings that can be applied across multiple worksheets. | Separate sections in a report sheet using a subtle top border combined with italicized labels.
// AI Prompts: Generate C# code using Aspose.Cells to apply bold font and a double bottom border to a range of cells. | Show how to build a reusable Style that includes background color, left and right borders, and font size, then apply it with StyleFlag. | Explain the difference between SetStyle(style) and SetStyle(style, flag) in Aspose.Cells and when to use each.

using System;
using System.Drawing;
using Aspose.Cells;

// Shows how to build a custom Style in Aspose.Cells, set the font to italic, add a thin black top border, use a StyleFlag to limit the changes to those two attributes, apply the style to cell A1, and save the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Create a custom style
        Style style = workbook.CreateStyle();

        // Set italic font
        style.Font.IsItalic = true;

        // Configure a thin top border (black color)
        style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
        style.Borders[BorderType.TopBorder].Color = Color.Black;

        // Define which style elements to apply
        StyleFlag flag = new StyleFlag();
        flag.FontItalic = true;   // Apply italic setting
        flag.TopBorder = true;    // Apply top border setting

        // Apply the style to cell A1
        Cell cell = cells["A1"];
        cell.PutValue("Italic with top border");
        cell.SetStyle(style, flag);

        // Save the workbook
        workbook.Save("CustomStyle.xlsx", SaveFormat.Xlsx);
    }
}
