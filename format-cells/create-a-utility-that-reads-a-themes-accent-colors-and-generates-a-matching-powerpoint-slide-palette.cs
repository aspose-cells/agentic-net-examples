// Title: C# Aspose.Cells Example – Extract Excel Theme Accent Colors and Create a Visual Palette Workbook
// Description: This C# sample uses Aspose.Cells for .NET to load an Excel workbook, read the six theme accent colors via GetThemeColor, and generate a new workbook that displays each accent label alongside a solid‑filled cell. The palette is saved as AccentPalette.xlsx and the RGB hex values are printed to the console. The code can be extended to apply the same colors to a PowerPoint slide master using Aspose.Slides.
// Keywords: Aspose.Cells | C# | .NET | Excel theme colors | GetThemeColor | Accent palette | Generate color palette workbook | PowerPoint slide colors | Aspose.Slides integration | GitHub example | open source code
// Common Searches: How to read Excel theme accent colors with Aspose.Cells C# | Create a color palette workbook from Excel theme using .NET | Aspose.Cells GetThemeColor example | Export Excel theme colors to PowerPoint slide master | C# code to display Excel theme colors in cells | GitHub Aspose.Cells theme color sample
// Developer Intent: Read the six accent colors defined in an Excel workbook’s theme and produce a new workbook that visualizes those colors for design reference or further use.
// Use Cases: Designers can quickly generate a reference palette that matches Excel theme colors for PowerPoint presentations. | Automated branding pipelines can extract theme colors and store them in a readable Excel file. | Quality‑control scripts can verify that a workbook uses the expected theme by comparing extracted hex values. | Developers can extend the sample to sync Excel theme colors with PowerPoint slide masters via Aspose.Slides.
// AI Prompts: Generate C# code using Aspose.Cells to retrieve ThemeColorType.Accent1‑Accent6 from a workbook and output their RGB hex values. | Modify the example to create a matching PowerPoint slide master palette with Aspose.Slides. | Add error handling that falls back to a default theme when the source file is missing or has no custom theme. | Provide a GitHub‑ready README that explains how to run the sample and integrate it into CI pipelines.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemeToPowerPoint
{
    // This C# sample uses Aspose.Cells for .NET to load an Excel workbook, read the six theme accent colors via GetThemeColor, and generate a new workbook that displays each accent label alongside a solid‑filled cell. The palette is saved as AccentPalette.xlsx and the RGB hex values are printed to the console. The code can be extended to apply the same colors to a PowerPoint slide master using Aspose.Slides.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel file that contains the theme.
                string excelPath = "SourceWithTheme.xlsx";

                // Load the workbook (create if file does not exist).
                Workbook sourceWorkbook;
                if (File.Exists(excelPath))
                {
                    sourceWorkbook = new Workbook(excelPath);
                }
                else
                {
                    sourceWorkbook = new Workbook(); // default theme
                }

                // Retrieve the six accent colors from the workbook's theme.
                Color[] accentColors = new Color[6];
                accentColors[0] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent1);
                accentColors[1] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent2);
                accentColors[2] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent3);
                accentColors[3] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent4);
                accentColors[4] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent5);
                accentColors[5] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent6);

                // Create a new workbook to display the accent palette.
                Workbook paletteWorkbook = new Workbook();
                Worksheet sheet = paletteWorkbook.Worksheets[0];
                sheet.Name = "Accent Palette";

                // Define layout parameters.
                const int startRow = 0;
                const int startColumn = 0;
                const int cellWidth = 20; // Approximate column width
                const int cellHeight = 30; // Approximate row height

                // Populate cells with accent colors and labels.
                for (int i = 0; i < accentColors.Length; i++)
                {
                    int row = startRow + i * 2; // Leave a blank row between entries

                    // Label cell.
                    Cell labelCell = sheet.Cells[row, startColumn];
                    labelCell.PutValue($"Accent{i + 1}");
                    // Optional: make label bold.
                    Style labelStyle = labelCell.GetStyle();
                    labelStyle.Font.IsBold = true;
                    labelCell.SetStyle(labelStyle);

                    // Color cell.
                    Cell colorCell = sheet.Cells[row + 1, startColumn];
                    // Apply background fill with the accent color.
                    Style style = colorCell.GetStyle();
                    style.ForegroundColor = accentColors[i];
                    style.Pattern = BackgroundType.Solid;
                    colorCell.SetStyle(style);
                    // Set a placeholder value to make the cell visible.
                    colorCell.PutValue(" ");

                    // Adjust column width and row height for better visibility.
                    sheet.Cells.SetColumnWidth(startColumn, cellWidth);
                    sheet.Cells.SetRowHeight(row, cellHeight);
                    sheet.Cells.SetRowHeight(row + 1, cellHeight);
                }

                // Save the generated Excel file.
                string palettePath = "AccentPalette.xlsx";
                paletteWorkbook.Save(palettePath, SaveFormat.Xlsx);

                Console.WriteLine("Accent colors extracted from Excel theme and saved to a palette workbook:");
                for (int i = 0; i < accentColors.Length; i++)
                {
                    Console.WriteLine($"Accent{i + 1}: #{accentColors[i].R:X2}{accentColors[i].G:X2}{accentColors[i].B:X2}");
                }
                Console.WriteLine($"Palette workbook saved to: {palettePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
