// Title: Create an Excel workbook that displays a WCAG contrast‑ratio matrix for the six default theme accent colors with colored headers and a red‑to‑green gradient using Aspose.Cells in C#
// AI Prompts: Write a C# program that uses Aspose.Cells to generate an .xlsx file containing a matrix of WCAG 2.0 contrast ratios for the six built‑in Excel theme accent colors, color the row and column headers with the corresponding accent, and shade each cell background from red (low contrast) to green (high contrast). | Extend the contrast‑matrix generator so it reads a JSON file with an arbitrary list of theme accent colors, rebuilds the matrix dynamically, and retains the same header coloring and value‑based gradient styling. | Add functionality to export the generated contrast‑ratio workbook to PDF while preserving the header background colors and the red‑to‑green gradient applied to the cells, using Aspose.Cells.
// Common Searches: how to calculate WCAG contrast ratio between Excel theme colors in C# with Aspose.Cells | generate a color contrast heatmap in an Excel file using Aspose.Cells C# | Aspose.Cells create matrix of theme accent colors with gradient background | C# program to output contrast ratio table for Excel accent colors to XLSX | save Aspose.Cells workbook as PDF with cell formatting intact
// Tags: Aspose.Cells generate contrast ratio matrix | C# compute WCAG luminance Aspose.Cells | Excel theme accent color heatmap Aspose.Cells | apply gradient background based on value Aspose.Cells | export workbook to PDF preserving styles Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ContrastMatrixApp
{
    // The example computes WCAG 2.0 luminance and contrast ratios for the six standard Excel theme accent colors, creates a new workbook, writes colored header rows and columns, fills a matrix with the contrast ratios, applies a red‑to‑green background gradient to each cell according to the ratio, auto‑fits columns, and saves the result as ThemeAccentContrastMatrix.xlsx.
    class ContrastMatrixGenerator
    {
        // Compute relative luminance of a color according to WCAG 2.0
        static double GetLuminance(Color color)
        {
            // Convert sRGB components to linear values
            double R = color.R / 255.0;
            double G = color.G / 255.0;
            double B = color.B / 255.0;

            R = (R <= 0.03928) ? R / 12.92 : Math.Pow((R + 0.055) / 1.055, 2.4);
            G = (G <= 0.03928) ? G / 12.92 : Math.Pow((G + 0.055) / 1.055, 2.4);
            B = (B <= 0.03928) ? B / 12.92 : Math.Pow((B + 0.055) / 1.055, 2.4);

            // Relative luminance formula
            return 0.2126 * R + 0.7152 * G + 0.0722 * B;
        }

        // Compute contrast ratio between two colors (WCAG 2.0)
        static double GetContrastRatio(Color c1, Color c2)
        {
            double L1 = GetLuminance(c1);
            double L2 = GetLuminance(c2);
            // Ensure L1 is the lighter color
            if (L1 < L2)
            {
                double temp = L1;
                L1 = L2;
                L2 = temp;
            }
            return Math.Round((L1 + 0.05) / (L2 + 0.05), 2);
        }

        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define a fallback set of six accent colors (standard Excel theme)
                Color[] accentColors = new Color[]
                {
                    Color.FromArgb(0, 112, 192),   // Accent 1
                    Color.FromArgb(255, 192, 0),   // Accent 2
                    Color.FromArgb(112, 48, 160),  // Accent 3
                    Color.FromArgb(0, 176, 80),    // Accent 4
                    Color.FromArgb(255, 0, 0),     // Accent 5
                    Color.FromArgb(0, 176, 240)    // Accent 6
                };

                // Add a worksheet to hold the matrix
                Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet.Name = "ContrastMatrix";

                // Write header row (accent colors)
                for (int col = 0; col < accentColors.Length; col++)
                {
                    Cell headerCell = sheet.Cells[0, col + 1];
                    headerCell.PutValue($"Accent {col + 1}");

                    Style style = headerCell.GetStyle();
                    style.ForegroundColor = accentColors[col];
                    style.Pattern = BackgroundType.Solid;
                    style.Font.Color = (GetLuminance(accentColors[col]) > 0.5) ? Color.Black : Color.White;
                    headerCell.SetStyle(style);
                }

                // Write header column (accent colors)
                for (int row = 0; row < accentColors.Length; row++)
                {
                    Cell headerCell = sheet.Cells[row + 1, 0];
                    headerCell.PutValue($"Accent {row + 1}");

                    Style style = headerCell.GetStyle();
                    style.ForegroundColor = accentColors[row];
                    style.Pattern = BackgroundType.Solid;
                    style.Font.Color = (GetLuminance(accentColors[row]) > 0.5) ? Color.Black : Color.White;
                    headerCell.SetStyle(style);
                }

                // Fill the matrix with contrast ratios
                for (int row = 0; row < accentColors.Length; row++)
                {
                    for (int col = 0; col < accentColors.Length; col++)
                    {
                        double ratio = GetContrastRatio(accentColors[row], accentColors[col]);
                        Cell cell = sheet.Cells[row + 1, col + 1];
                        cell.PutValue(ratio);

                        // Map ratio 1–21 to a gradient from red (low) to green (high)
                        int red = (int)(255 - Math.Min(255, (ratio - 1) * 12));
                        int green = (int)Math.Min(255, (ratio - 1) * 12);
                        Color bgColor = Color.FromArgb(red, green, 0);

                        Style style = cell.GetStyle();
                        style.ForegroundColor = bgColor;
                        style.Pattern = BackgroundType.Solid;
                        style.Font.Color = (GetLuminance(bgColor) > 0.5) ? Color.Black : Color.White;
                        cell.SetStyle(style);
                    }
                }

                // Auto-fit columns for better appearance
                sheet.AutoFitColumns();

                // Define output file path
                string outputPath = "ThemeAccentContrastMatrix.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
