// Title: Create an Excel workbook with a custom theme and apply accent colors to header and alternating rows using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to generate a new workbook, define a dictionary of ThemeColorType accent colors, and style the first row as a header with those colors. | Provide a C# example that applies alternating row background colors based on custom theme accents in an Aspose.Cells worksheet and saves the file.
// Common Searches: Aspose.Cells C# how to set custom theme accent colors for a worksheet | C# apply header style with custom theme colors using Aspose.Cells | Aspose.Cells alternating row shading with custom theme colors in .NET | Create Excel file with custom theme colors programmatically using Aspose.Cells
// Tags: custom theme accent colors Aspose.Cells .NET | header formatting using ThemeColorType Aspose.Cells | alternating row background Aspose.Cells workbook | dictionary of ThemeColorType to Color C# | save workbook as .xlsx Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, defines a set of custom ThemeColorType accent colors in a dictionary, writes sample category/value data, styles the header row with the first accent color, applies alternating row background colors using other accents, and saves the result as CustomThemeSample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define custom theme colors using a dictionary (Aspose.Cells Theme class not available in this version)
            var customColors = new Dictionary<ThemeColorType, Color>
            {
                { ThemeColorType.Accent1, Color.FromArgb(91, 155, 213) },   // Light blue
                { ThemeColorType.Accent2, Color.FromArgb(237, 125, 49) },   // Orange
                { ThemeColorType.Accent3, Color.FromArgb(165, 165, 165) }, // Gray
                { ThemeColorType.Accent4, Color.FromArgb(255, 192, 0) },   // Gold
                { ThemeColorType.Accent5, Color.FromArgb(112, 173, 71) },  // Green
                { ThemeColorType.Accent6, Color.FromArgb(68, 114, 196) }   // Dark blue
            };

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            string[] categories = { "A", "B", "C", "D", "E" };
            double[] values = { 10, 20, 30, 40, 50 };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue(values[i]);   // Column B
            }

            // Apply theme colors to header style
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = customColors[ThemeColorType.Accent1];
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = customColors[ThemeColorType.Accent2];
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the header style
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange("A1:B1");
            headerRange.ApplyStyle(headerStyle, new StyleFlag { Font = true, CellShading = true });

            // Apply alternating row colors using theme accents
            for (int row = 2; row <= categories.Length + 1; row++)
            {
                Style rowStyle = workbook.CreateStyle();
                Color bgColor = (row % 2 == 0)
                    ? customColors[ThemeColorType.Accent3]
                    : customColors[ThemeColorType.Accent4];
                rowStyle.ForegroundColor = bgColor;
                rowStyle.Pattern = BackgroundType.Solid;

                Aspose.Cells.Range dataRange = sheet.Cells.CreateRange(row - 1, 0, 1, 2);
                dataRange.ApplyStyle(rowStyle, new StyleFlag { CellShading = true });
            }

            // Save the workbook
            workbook.Save("CustomThemeSample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
