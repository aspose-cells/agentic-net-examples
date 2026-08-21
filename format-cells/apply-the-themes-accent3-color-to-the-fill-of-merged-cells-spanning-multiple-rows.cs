// Title: Aspose.Cells for .NET: Apply Theme Accent3 Fill to a Multi‑Row Merged Range (C#)
// Description: C# sample that creates a workbook, merges a 5‑row × 3‑column range, and applies a solid fill using the workbook’s ThemeColor Accent3. The styled range is saved as an Excel file.
// Keywords: Aspose.Cells C# merge cells | ThemeColor Accent3 fill | apply theme color to merged range | Excel solid fill Aspose.Cells | StyleFlag CellShading example | Workbook theme color .NET | Aspose.Cells style merged cells | set background using ThemeColorType | Excel header merged cells theme | Aspose.Cells sample code
// Common Searches: Aspose.Cells how to fill merged cells with theme Accent3 | C# apply workbook theme color to merged range Aspose.Cells | Set solid fill for merged cells using ThemeColor in .NET | Merge cells and style with Accent3 color in Aspose.Cells | Example code for ThemeColorType.Accent3 with merged cells | Aspose.Cells StyleFlag CellShading usage | Create multi‑row header with theme color in Excel using Aspose | GitHub Aspose.Cells merged cells fill example
// Developer Intent: Apply the workbook’s Accent3 theme color as a solid fill to a merged range that spans multiple rows.
// Use Cases: Design a multi‑row header in a report that matches the workbook’s theme. | Highlight a merged block in a dashboard with the Accent3 color for visual emphasis. | Generate Excel templates where merged sections inherit the theme’s accent color automatically. | Create printable invoices where the merged title area uses the theme’s Accent3 fill.
// AI Prompts: Generate C# code that merges a range of cells and applies the workbook’s Accent3 theme color as a solid background using Aspose.Cells. | Explain how to use ThemeColorType.Accent3 with StyleFlag.CellShading to style merged cells in Aspose.Cells for .NET. | Show an example of changing the theme color index for a merged cell style in Aspose.Cells. | Provide a step‑by‑step guide to apply a gradient fill based on a theme color to a merged range in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// C# sample that creates a workbook, merges a 5‑row × 3‑column range, and applies a solid fill using the workbook’s ThemeColor Accent3. The styled range is saved as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a range that spans multiple rows (5 rows x 3 columns) and merge it
            AsposeRange mergedRange = sheet.Cells.CreateRange(0, 0, 5, 3);
            mergedRange.Merge();

            // Add some text to the merged cell for visual reference
            sheet.Cells[0, 0].PutValue("Merged with Accent3 Fill");

            // Create a style and set solid fill pattern
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;

            // Apply the theme's Accent3 color to the fill using ThemeColor
            style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);

            // Specify that cell shading (fill) should be applied
            StyleFlag flag = new StyleFlag();
            flag.CellShading = true;

            // Apply the style to the merged range
            mergedRange.ApplyStyle(style, flag);

            // Save the workbook (lifecycle rule)
            workbook.Save("MergedAccent3.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
