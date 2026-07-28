// Title: C# – Export Workbook Theme Palette to Excel with Aspose.Cells
// Description: Demonstrates how to create or load a workbook, add a "ThemePalette" worksheet, list every ThemeColorType with its ARGB hex value, show a visual sample cell, auto‑fit columns, and save the file as ThemePaletteExport.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# export theme palette | list workbook theme colors | ThemeColorType ARGB values | generate theme palette report | .NET Excel theme colors | sample cell color styling | auto‑fit columns Aspose | save workbook as xlsx
// Common Searches: export theme colors from Aspose.Cells C# | how to list workbook theme palette in Excel | Aspose.Cells get theme color ARGB | create theme palette worksheet Aspose | C# code to export Excel theme colors
// Developer Intent: Produce an Excel sheet that enumerates all theme color types, their ARGB hex codes, and a visual sample, enabling designers to review the workbook’s color scheme.
// Use Cases: Provide designers with a quick reference of a workbook’s theme colors for branding decisions. | Generate diagnostic reports that compare custom theme colors against the default palette. | Automate documentation of theme palettes across multiple workbooks in a CI pipeline.
// AI Prompts: Write C# code with Aspose.Cells to export the workbook’s theme palette, including ARGB values and a sample cell for each color. | Extend the example to add separate columns for Red, Green, and Blue components of each theme color. | Add conditional formatting that highlights theme colors differing from the default Aspose.Cells theme.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemePaletteExport
{
    // Demonstrates how to create or load a workbook, add a "ThemePalette" worksheet, list every ThemeColorType with its ARGB hex value, show a visual sample cell, auto‑fit columns, and save the file as ThemePaletteExport.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Optionally modify some theme colors for demonstration
            // workbook.SetThemeColor(ThemeColorType.Accent1, Color.Orange);
            // workbook.SetThemeColor(ThemeColorType.Accent2, Color.Purple);

            // Add a new worksheet to hold the theme palette information
            Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
            sheet.Name = "ThemePalette";

            // Write header row
            sheet.Cells["A1"].PutValue("Theme Color Type");
            sheet.Cells["B1"].PutValue("ARGB Value");
            sheet.Cells["C1"].PutValue("Sample");

            // Apply bold style to header
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Size = 12;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.ForegroundColor = Color.LightGray;
            sheet.Cells.CreateRange("A1:C1").SetStyle(headerStyle);

            // Iterate through all ThemeColorType values (0 to 11)
            int rowIndex = 1; // zero‑based index; row 2 in Excel
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the actual color for the current theme type
                Color themeColor = workbook.GetThemeColor(type);

                // Write the theme type name
                sheet.Cells[rowIndex, 0].PutValue(type.ToString());

                // Write the ARGB value as a hex string
                string argbHex = $"0x{themeColor.ToArgb():X8}";
                sheet.Cells[rowIndex, 1].PutValue(argbHex);

                // Create a style that uses the theme color as foreground
                Style sampleStyle = workbook.CreateStyle();
                sampleStyle.ForegroundColor = themeColor;
                sampleStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the sample cell and put a label
                Cell sampleCell = sheet.Cells[rowIndex, 2];
                sampleCell.PutValue("Sample");
                sampleCell.SetStyle(sampleStyle);

                rowIndex++;
            }

            // Auto‑fit columns for better readability
            sheet.AutoFitColumns();

            // Save the workbook containing the exported theme palette
            workbook.Save("ThemePaletteExport.xlsx");
        }
    }
}
