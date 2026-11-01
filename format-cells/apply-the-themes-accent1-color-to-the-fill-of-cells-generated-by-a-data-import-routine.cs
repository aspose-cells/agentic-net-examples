// Title: How to fill a cell range with the workbook's Theme Accent1 color using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set background of a range to the workbook's Accent1 theme color | Apply theme accent color as fill to imported data cells in Aspose.Cells | Retrieve ThemeColorType.Accent1 and use it for cell shading in .NET | Create a style with solid fill using workbook theme color in Aspose.Cells | Use StyleFlag to apply only fill to a range in Aspose.Cells
// Tags: apply theme accent color fill Aspose.Cells | style flag cell shading .NET | create style with workbook theme color C# | range fill using ThemeColorType.Accent1 | solid background fill Excel file Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, populates cells A1:C3, retrieves the workbook's Theme Accent1 color (fallback to LightBlue), builds a style with a solid fill, applies only the fill to the range using a StyleFlag, and saves the file as an XLSX document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill cells A1:C3 with sample data
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a style that uses the workbook's theme Accent1 color
            Style accentStyle = workbook.CreateStyle();

            // Retrieve Accent1 color from the theme; fallback to LightBlue if unavailable
            Color accentColor = workbook.GetThemeColor(ThemeColorType.Accent1);
            if (accentColor == Color.Empty)
                accentColor = Color.LightBlue;

            accentStyle.ForegroundColor = accentColor;
            accentStyle.Pattern = BackgroundType.Solid; // Solid fill

            // Apply only the fill (shading) part of the style
            StyleFlag flag = new StyleFlag { CellShading = true };

            // Apply the style to the range A1:C3
            AsposeRange importedRange = sheet.Cells.CreateRange(0, 0, 3, 3);
            importedRange.ApplyStyle(accentStyle, flag);

            // Save the workbook
            string outputPath = "ImportedData_With_Accent1_Fill.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
