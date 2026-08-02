// Title: Set Column Font Size, Color, and Underline with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, select column B, build a Style with a 14‑point blue font and single underline, enable the matching StyleFlag properties, apply the style to the entire column, and save the file as ColumnFontSettings.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# column font | set column font size Aspose.Cells | column font color Aspose.Cells | underline column text Aspose.Cells | .NET Excel column formatting | Apply StyleFlag to column | Excel column style programmatically | Aspose.Cells column formatting example | C# Excel column font settings | Aspose.Cells workbook column style
// Common Searches: Aspose.Cells change font size for a whole column | C# apply font color to Excel column using Aspose.Cells | How to underline text in an Excel column with Aspose.Cells .NET | Set column style with StyleFlag Aspose.Cells | Apply same font to all cells in a column C# Aspose
// Developer Intent: Apply uniform font formatting (size, color, underline) to every cell in a specific worksheet column.
// Use Cases: Highlight a header column in generated reports with blue, 14‑point, underlined text for visual emphasis. | Enforce corporate spreadsheet style by programmatically applying the same font settings to target columns across multiple workbooks. | Provide end‑users a way to choose column appearance (size, color, underline) before exporting data to Excel.
// AI Prompts: Generate C# code using Aspose.Cells to set column C to a 12‑point red double‑underlined font. | Show how to apply distinct font styles (size, color, underline) to columns A, B, and D in a single workbook with Aspose.Cells. | Create a reusable method that accepts font size, color, underline type, and column index, then applies the style using StyleFlag in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

// Shows how to create a workbook, select column B, build a Style with a 14‑point blue font and single underline, enable the matching StyleFlag properties, apply the style to the entire column, and save the file as ColumnFontSettings.xlsx using Aspose.Cells for .NET.
class ConfigureColumnFont
{
    static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the desired column (e.g., column B, index 1)
        Column column = worksheet.Cells.Columns[1];

        // Create a style and configure the font properties
        Style style = workbook.CreateStyle();
        style.Font.Size = 14;                                 // Font size
        style.Font.Color = Color.Blue;                        // Font color
        style.Font.Underline = FontUnderlineType.Single;      // Underline

        // Define which font attributes should be applied
        StyleFlag flag = new StyleFlag
        {
            FontSize = true,
            FontColor = true,
            FontUnderline = true
        };

        // Apply the style to the entire column
        column.ApplyStyle(style, flag);

        // Save the workbook
        string outputPath = "ColumnFontSettings.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
