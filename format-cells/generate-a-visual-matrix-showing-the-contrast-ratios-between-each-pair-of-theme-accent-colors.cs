// Title: Generate an Excel contrast‑ratio matrix for theme accent colors with Aspose.Cells in C#
// Common Searches: how to calculate WCAG contrast ratio between two colors using Aspose.Cells C# | create a contrast matrix of theme accent colors in an Excel file with Aspose.Cells | C# Aspose.Cells example to fill a worksheet with color contrast values | generate Excel sheet showing contrast ratios for multiple theme colors using .NET | Aspose.Cells compute relative luminance and contrast ratio for accent colors
// Tags: compute contrast ratios using Aspose.Cells | write color contrast matrix to Excel workbook | WCAG relative luminance calculation C# | theme accent colors contrast analysis .NET | export contrast ratio table to .xlsx with Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ContrastMatrixDemo
{
    // The sample creates a new workbook, defines six theme accent colors, adds row and column headers, computes WCAG contrast ratios for every color pair using relative luminance, writes the rounded ratios into the worksheet to form a matrix, and saves the result as ContrastMatrix.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define a set of accent colors (fallback if ThemeColors is unavailable)
                Color[] accentColors = new Color[]
                {
                    Color.FromArgb(0, 120, 215),   // Accent 1
                    Color.FromArgb(232, 17, 35),   // Accent 2
                    Color.FromArgb(255, 185, 0),   // Accent 3
                    Color.FromArgb(0, 204, 106),   // Accent 4
                    Color.FromArgb(255, 0, 255),   // Accent 5
                    Color.FromArgb(0, 0, 0)        // Accent 6
                };

                int accentCount = accentColors.Length;

                // Write header labels (accent indices) in the first row and column
                for (int i = 0; i < accentCount; i++)
                {
                    // Column headers (starting from B1)
                    cells[0, i + 1].PutValue($"Accent {i + 1}");
                    // Row headers (starting from A2)
                    cells[i + 1, 0].PutValue($"Accent {i + 1}");
                }

                // Compute contrast ratios for each pair and fill the matrix
                for (int row = 0; row < accentCount; row++)
                {
                    for (int col = 0; col < accentCount; col++)
                    {
                        double ratio = ComputeContrastRatio(accentColors[row], accentColors[col]);
                        // Store the ratio with two decimal places
                        cells[row + 1, col + 1].PutValue(Math.Round(ratio, 2));
                    }
                }

                // Save the workbook to the current directory
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ContrastMatrix.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Calculates the contrast ratio between two colors according to WCAG.
        static double ComputeContrastRatio(Color c1, Color c2)
        {
            double l1 = GetRelativeLuminance(c1);
            double l2 = GetRelativeLuminance(c2);

            // Ensure l1 is the lighter color
            if (l2 > l1)
            {
                double temp = l1;
                l1 = l2;
                l2 = temp;
            }

            return (l1 + 0.05) / (l2 + 0.05);
        }

        // Calculates the relative luminance of a color according to WCAG.
        static double GetRelativeLuminance(Color color)
        {
            // Convert sRGB components to linear values
            double RsRGB = color.R / 255.0;
            double GsRGB = color.G / 255.0;
            double BsRGB = color.B / 255.0;

            double R = (RsRGB <= 0.03928) ? RsRGB / 12.92 : Math.Pow((RsRGB + 0.055) / 1.055, 2.4);
            double G = (GsRGB <= 0.03928) ? GsRGB / 12.92 : Math.Pow((GsRGB + 0.055) / 1.055, 2.4);
            double B = (BsRGB <= 0.03928) ? BsRGB / 12.92 : Math.Pow((BsRGB + 0.055) / 1.055, 2.4);

            // Apply the luminance coefficients
            return 0.2126 * R + 0.7152 * G + 0.0722 * B;
        }
    }
}
