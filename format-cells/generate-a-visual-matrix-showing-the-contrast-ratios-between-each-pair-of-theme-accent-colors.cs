// Title: Generate a Theme Accent Contrast Ratio Matrix in Excel with Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, extracts the six Excel theme accent colors, calculates the WCAG contrast ratio for every color pair, and writes the values into a formatted matrix with bold headers. Ratios are rounded to two decimals and the file is saved as ThemeAccentContrastMatrix.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel theme colors | contrast ratio | WCAG accessibility | theme accent matrix | GetThemeColor | ThemeColorType | relative luminance | color contrast analysis
// Common Searches: Aspose.Cells calculate contrast ratio between theme accent colors | C# generate Excel matrix of theme color contrast | WCAG contrast matrix for Excel theme colors using Aspose | How to compute color contrast in Aspose.Cells .NET
// Developer Intent: Create an Excel sheet that lists WCAG contrast ratios for all pairs of workbook theme accent colors.
// Use Cases: Check whether a workbook’s theme colors satisfy accessibility contrast standards before distribution. | Provide designers with a quick reference to select accent colors that achieve sufficient contrast. | Integrate color‑contrast validation into automated build or CI pipelines for multiple Excel files.
// AI Prompts: Highlight cells in red when the contrast ratio is below 4.5 to flag non‑compliant pairs. | Add a summary row that counts how many color pairs fail the WCAG AA threshold. | Export the generated contrast matrix to a CSV file in addition to the Excel workbook.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeContrastMatrix
{
    // This example creates a new workbook, extracts the six Excel theme accent colors, calculates the WCAG contrast ratio for every color pair, and writes the values into a formatted matrix with bold headers. Ratios are rounded to two decimals and the file is saved as ThemeAccentContrastMatrix.xlsx.
    public class ContrastMatrixGenerator
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define the six accent theme color types
                ThemeColorType[] accents = new ThemeColorType[]
                {
                    ThemeColorType.Accent1,
                    ThemeColorType.Accent2,
                    ThemeColorType.Accent3,
                    ThemeColorType.Accent4,
                    ThemeColorType.Accent5,
                    ThemeColorType.Accent6
                };

                // Write header row (column titles)
                for (int col = 0; col < accents.Length; col++)
                {
                    cells[0, col + 1].PutValue(accents[col].ToString());
                    cells[0, col + 1].SetStyle(CreateHeaderStyle(workbook));
                }

                // Write header column (row titles)
                for (int row = 0; row < accents.Length; row++)
                {
                    cells[row + 1, 0].PutValue(accents[row].ToString());
                    cells[row + 1, 0].SetStyle(CreateHeaderStyle(workbook));
                }

                // Compute contrast ratios for each pair and fill the matrix
                for (int i = 0; i < accents.Length; i++)
                {
                    Color colorA = workbook.GetThemeColor(accents[i]);

                    for (int j = 0; j < accents.Length; j++)
                    {
                        Color colorB = workbook.GetThemeColor(accents[j]);

                        double ratio = CalculateContrastRatio(colorA, colorB);

                        // Write the ratio with two decimal places
                        Cell cell = cells[i + 1, j + 1];
                        cell.PutValue(Math.Round(ratio, 2));

                        // Apply number format style
                        cell.SetStyle(CreateNumberStyle(workbook));
                    }
                }

                // Adjust column widths for readability
                sheet.AutoFitColumns();

                // Save the workbook
                workbook.Save("ThemeAccentContrastMatrix.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while generating the contrast matrix: {ex.Message}");
            }
        }

        // Creates a simple style for header cells (bold, centered)
        private static Style CreateHeaderStyle(Workbook wb)
        {
            Style style = wb.CreateStyle();
            style.Font.IsBold = true;
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            return style;
        }

        // Creates a style that formats numbers with two decimal places
        private static Style CreateNumberStyle(Workbook wb)
        {
            Style style = wb.CreateStyle();
            style.Custom = "0.00";
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            return style;
        }

        // Calculates the WCAG contrast ratio between two colors
        private static double CalculateContrastRatio(Color c1, Color c2)
        {
            double L1 = GetRelativeLuminance(c1);
            double L2 = GetRelativeLuminance(c2);

            // Ensure L1 is the lighter luminance
            if (L1 < L2)
            {
                double temp = L1;
                L1 = L2;
                L2 = temp;
            }

            return (L1 + 0.05) / (L2 + 0.05);
        }

        // Computes the relative luminance of a color per WCAG definition
        private static double GetRelativeLuminance(Color color)
        {
            // Convert sRGB components to linear values (0..1)
            double RsRGB = color.R / 255.0;
            double GsRGB = color.G / 255.0;
            double BsRGB = color.B / 255.0;

            double R = (RsRGB <= 0.03928) ? RsRGB / 12.92 : Math.Pow((RsRGB + 0.055) / 1.055, 2.4);
            double G = (GsRGB <= 0.03928) ? GsRGB / 12.92 : Math.Pow((GsRGB + 0.055) / 1.055, 2.4);
            double B = (BsRGB <= 0.03928) ? BsRGB / 12.92 : Math.Pow((BsRGB + 0.055) / 1.055, 2.4);

            // Relative luminance formula
            return 0.2126 * R + 0.7152 * G + 0.0722 * B;
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ContrastMatrixGenerator.Run();
        }
    }
}
