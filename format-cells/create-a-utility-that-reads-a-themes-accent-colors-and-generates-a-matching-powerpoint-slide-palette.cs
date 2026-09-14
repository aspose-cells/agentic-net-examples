// Title: Generate an Excel palette worksheet that visualizes a workbook’s six theme accent colors with Aspose.Cells for .NET (C#)
// AI Prompts: Create a C# method using Aspose.Cells that opens an .xlsx file, accesses the workbook’s ThemeColorScheme, and returns the six Accent colors as a List<Color>. | Write C# code that adds a new worksheet named "Accent Palette" to an existing workbook, fills separate cells with each accent color as a solid background, labels them "Accent 1"‑"Accent 6", and saves the file as Xlsx.
// Common Searches: aspnet how to extract theme accent colors from an Excel file using Aspose.Cells | c# create a color palette sheet from Excel theme with Aspose.Cells | aspose.cells retrieve six accent colors and display them in a new worksheet | generate visual theme palette in Excel programmatically with Aspose.Cells C# | save workbook with added accent color palette using Aspose.Cells .NET
// Tags: extract workbook theme accent colors Aspose.Cells | add palette worksheet Aspose.Cells C# | solid background fill cell Aspose.Cells | save workbook as Xlsx Aspose.Cells | dynamic theme access Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemePaletteUtility
{
    // The PaletteGenerator class uses Aspose.Cells to load an Excel workbook, dynamically accesses its ThemeColorScheme to collect the six accent colors, adds a new worksheet named "Accent Palette", configures column widths and row heights, applies each accent as a solid background fill to individual cells, labels the cells with "Accent 1" through "Accent 6", ensures the output directory exists, and saves the modified workbook as an Xlsx file.
    public class PaletteGenerator
    {
        /// <param name="workbookPath">Full path to the Excel file.</param>
        /// <returns>List of accent colors (Accent1‑Accent6).</returns>
        public List<Color> GetAccentColors(string workbookPath)
        {
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook not found: {workbookPath}");

            try
            {
                // Load the workbook
                var wb = new Workbook(workbookPath);

                // Use dynamic to access theme information (avoids compile‑time dependency on Theme classes)
                dynamic theme = wb.Theme;
                if (theme == null)
                    throw new InvalidOperationException("The workbook does not contain a theme.");

                dynamic scheme = theme.ThemeColorScheme;
                if (scheme == null)
                    throw new InvalidOperationException("Theme color scheme is unavailable.");

                // Collect the six accent colors
                return new List<Color>
                {
                    (Color)scheme.Accent1,
                    (Color)scheme.Accent2,
                    (Color)scheme.Accent3,
                    (Color)scheme.Accent4,
                    (Color)scheme.Accent5,
                    (Color)scheme.Accent6
                };
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to retrieve accent colors.", ex);
            }
        }

        /// <param name="sourcePath">Path to the source workbook.</param>
        /// <param name="outputPath">Path where the new workbook will be saved.</param>
        public void CreatePaletteSheet(string sourcePath, string outputPath)
        {
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

            try
            {
                // Load the source workbook
                var wb = new Workbook(sourcePath);

                // Retrieve accent colors
                List<Color> accents = GetAccentColors(sourcePath);

                // Add a new worksheet for the palette
                int sheetIndex = wb.Worksheets.Add();
                Worksheet paletteSheet = wb.Worksheets[sheetIndex];
                paletteSheet.Name = "Accent Palette";

                // Layout parameters
                const int startRow = 0;
                const int startColumn = 0;
                const int boxWidth = 15;   // column width in characters
                const int boxHeight = 30;  // row height in points (adjusted for visibility)

                // Create a cell for each accent color
                for (int i = 0; i < accents.Count; i++)
                {
                    int col = startColumn + i * 2; // leave a column gap between boxes

                    // Set column width and row height for the color box
                    paletteSheet.Cells.SetColumnWidth(col, boxWidth);
                    paletteSheet.Cells.SetRowHeight(startRow, boxHeight);

                    // Apply background color to the cell
                    Style style = wb.CreateStyle();
                    style.ForegroundColor = accents[i];
                    style.Pattern = BackgroundType.Solid;
                    paletteSheet.Cells[startRow, col].SetStyle(style);

                    // Add label below the color box
                    int labelRow = startRow + 1;
                    paletteSheet.Cells.SetRowHeight(labelRow, 12);
                    paletteSheet.Cells[labelRow, col].PutValue($"Accent {i + 1}");
                    Style labelStyle = wb.CreateStyle();
                    labelStyle.HorizontalAlignment = TextAlignmentType.Center;
                    paletteSheet.Cells[labelRow, col].SetStyle(labelStyle);
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                wb.Save(outputPath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to create palette sheet.", ex);
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            try
            {
                string sourceXlsx = @"C:\Docs\SourceWorkbook.xlsx";
                string outputXlsx = @"C:\Docs\PaletteWorkbook.xlsx";

                var generator = new PaletteGenerator();
                generator.CreatePaletteSheet(sourceXlsx, outputXlsx);

                Console.WriteLine("Palette worksheet created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
