// Title: Use Aspose.Cells for .NET to fill a merged cell range with the workbook's Accent3 theme color
// AI Prompts: Merge a range of cells (e.g., B2:D5) and set its background to the workbook's Accent3 theme color using Aspose.Cells in C#. | Retrieve the Accent3 color from the workbook theme with GetThemeColor and apply a solid fill style to the top‑left cell of the merged block. | Create a Style with BackgroundType.Solid, assign the Accent3 color, and apply it to a merged range in an Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to apply a theme accent color to merged cells | set background color of merged range using GetThemeColor Aspose.Cells | fill merged cells with workbook theme Accent3 in .NET | apply solid fill style to merged cells Aspose.Cells example | retrieve theme colors and style merged cells in Excel with Aspose
// Tags: merge cells and apply theme accent fill Aspose.Cells | GetThemeColor Accent3 solid style .NET | style merged range background Aspose.Cells | theme color fill for merged cells C# | Aspose.Cells workbook theme color usage

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a workbook, merges cells B2:D5, obtains the workbook's Accent3 theme color via GetThemeColor, builds a solid fill style with that color, applies the style to the top‑left cell of the merged range, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to merge (e.g., B2:D5)
            int startRow = 1;      // Row 2 (zero‑based)
            int startColumn = 1;   // Column B
            int rowCount = 4;      // Rows 2‑5
            int columnCount = 3;   // Columns B‑D

            // Merge the cells spanning multiple rows
            sheet.Cells.Merge(startRow, startColumn, rowCount, columnCount);

            // Retrieve the theme's Accent3 color using GetThemeColor
            Color accent3 = workbook.GetThemeColor(ThemeColorType.Accent3);

            // Create a style with solid fill using Accent3
            Style style = workbook.CreateStyle();
            style.ForegroundColor = accent3;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the merged range (top‑left cell is sufficient)
            StyleFlag flag = new StyleFlag { All = true };
            sheet.Cells[startRow, startColumn].SetStyle(style, flag);

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
