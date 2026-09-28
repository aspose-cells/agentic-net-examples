// Title: How to set font size, color, and underline for an entire column with Aspose.Cells for .NET (C#)
// AI Prompts: Create a 14‑point dark‑green font with single underline, assign it to a Style, and apply only the font attributes to column B of a new workbook using Aspose.Cells in C#. | Use a StyleFlag to restrict a custom style to font properties and apply that style to a specific worksheet column programmatically.
// Common Searches: Aspose.Cells C# set column font size and color without affecting other formatting | Apply underline to an entire column in an Excel file using Aspose.Cells for .NET | How to use StyleFlag to change only font attributes of a column with Aspose.Cells
// Tags: column font styling Aspose.Cells C# | apply dark green underlined font to column Aspose.Cells | StyleFlag apply font only Aspose.Cells | custom column style Aspose.Cells workbook

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new workbook, defines a style with a 14‑point dark‑green underlined font, uses a StyleFlag to limit the style to font attributes, applies the style to column B, and saves the file as ColumnFontSettings.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style object for the column
            Style columnStyle = workbook.CreateStyle();

            // Set the desired font size
            columnStyle.Font.Size = 14;

            // Set the desired font color
            columnStyle.Font.Color = Color.DarkGreen;

            // Enable underline for the font
            columnStyle.Font.Underline = FontUnderlineType.Single;

            // Apply only font attributes
            StyleFlag flag = new StyleFlag { Font = true };

            // Apply the style to column B (index 1)
            sheet.Cells.Columns[1].ApplyStyle(columnStyle, flag);

            // Save the workbook to a file
            string outputPath = "ColumnFontSettings.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
