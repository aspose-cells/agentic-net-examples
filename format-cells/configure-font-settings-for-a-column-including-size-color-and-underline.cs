// Title: How to set font size, color, and underline for an entire column with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set column font size and color | apply single underline to a whole column using Aspose.Cells .NET | use StyleFlag to change only font attributes of a column in Aspose.Cells | C# example for formatting column B font properties with Aspose.Cells
// Tags: column font styling Aspose.Cells | set font size color underline Aspose.Cells C# | StyleFlag apply font only Aspose.Cells | apply style to column Aspose.Cells .NET | column formatting workbook Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// Creates a style with a 14‑point dark red font and single underline, applies it to column B using a StyleFlag that limits changes to font attributes, and saves the workbook as ColumnFontSettings.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Specify the column index (e.g., column B = index 1)
        int columnIndex = 1;

        // Create a style object to configure font settings
        Style style = workbook.CreateStyle();

        // Set font size
        style.Font.Size = 14;

        // Set font color
        style.Font.Color = Color.DarkRed;

        // Set underline style
        style.Font.Underline = FontUnderlineType.Single;

        // Apply the style to the entire column (only font-related attributes)
        StyleFlag flag = new StyleFlag();
        flag.Font = true; // apply only font changes
        sheet.Cells.Columns[columnIndex].ApplyStyle(style, flag);

        // Save the workbook
        workbook.Save("ColumnFontSettings.xlsx");
    }
}
