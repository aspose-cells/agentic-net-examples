// Title: Create a Theme Accent Contrast‑Ratio Matrix in Excel with Aspose.Cells for .NET
// Description: C# example that creates a workbook, extracts the six theme accent colors using GetThemeColor, computes WCAG luminance and contrast ratios for each color pair, writes the results into a styled matrix with row/column headers, auto‑fits columns, and saves the file as ThemeAccentContrastMatrix.xlsx.
// Keywords: Aspose.Cells | C# | .NET | contrast ratio | WCAG | theme accent colors | Excel matrix | GetThemeColor | accessibility | color contrast calculation
// Common Searches: Aspose.Cells compute contrast ratio between theme colors | C# generate contrast matrix for Excel theme accents | how to calculate WCAG contrast in Aspose.Cells | Excel theme accent color contrast table example
// Developer Intent: Generate an Excel worksheet that lists WCAG contrast ratios for every combination of the workbook’s theme accent colors.
// Use Cases: Verify that a workbook’s default accent palette complies with accessibility contrast standards. | Provide designers with a numeric view of color relationships for branding decisions. | Automate contrast‑ratio reporting across multiple workbooks to maintain consistent visual quality.
// AI Prompts: Add conditional formatting to the matrix that highlights ratios below 4.5 with a red fill using Aspose.Cells. | Extend the program to export the contrast‑ratio matrix to CSV while preserving header labels. | Include the theme’s background and text colors in the matrix and flag any pairs that fail WCAG AA requirements.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeContrastMatrix
{
    // C# example that creates a workbook, extracts the six theme accent colors using GetThemeColor, computes WCAG luminance and contrast ratios for each color pair, writes the results into a styled matrix with row/column headers, auto‑fits columns, and saves the file as ThemeAccentContrastMatrix.xlsx.
    class Program
    {
        // Compute relative luminance of a color according to WCAG definition
        static double GetLuminance(Color color)
        {
            // Convert RGB components to linear sRGB (0..1)
            double R = color.R / 255.0;
            double G = color.G / 255.0;
            double B = color.B / 255.0;

            double Linear(double channel)
            {
                return (channel <= 0.03928) ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);
            }

            double r = Linear(R);
            double g = Linear(G);
            double b = Linear(B);

            // Relative luminance
            return 0.2126 * r + 0.7152 * g + 0.0722 * b;
        }

        // Compute contrast ratio between two colors
        static double GetContrastRatio(Color c1, Color c2)
        {
            double L1 = GetLuminance(c1);
            double L2 = GetLuminance(c2);
            // Ensure L1 is the lighter color
            if (L2 > L1) (L1, L2) = (L2, L1);
            return (L1 + 0.05) / (L2 + 0.05);
        }

        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the six accent theme colors
            ThemeColorType[] accentTypes = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            Color[] accentColors = new Color[accentTypes.Length];
            for (int i = 0; i < accentTypes.Length; i++)
            {
                accentColors[i] = workbook.GetThemeColor(accentTypes[i]);
            }

            // Write headers (Accent names)
            for (int i = 0; i < accentTypes.Length; i++)
            {
                // Column headers (row 2, starting from column C)
                sheet.Cells[1, i + 2].PutValue(accentTypes[i].ToString());
                // Row headers (column B, starting from row 3)
                sheet.Cells[i + 2, 1].PutValue(accentTypes[i].ToString());
            }

            // Compute and fill contrast ratios
            for (int row = 0; row < accentColors.Length; row++)
            {
                for (int col = 0; col < accentColors.Length; col++)
                {
                    double ratio = GetContrastRatio(accentColors[row], accentColors[col]);
                    // Round to two decimal places for readability
                    double rounded = Math.Round(ratio, 2);
                    // Place value in matrix (starting at cell C3)
                    sheet.Cells[row + 2, col + 2].PutValue(rounded);
                }
            }

            // Optional: format the matrix for better visual appearance
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.ForegroundColor = Color.LightGray;

            // Apply header style to top row and left column
            for (int i = 0; i < accentTypes.Length; i++)
            {
                sheet.Cells[1, i + 2].SetStyle(headerStyle); // column headers
                sheet.Cells[i + 2, 1].SetStyle(headerStyle); // row headers
            }

            // Auto-fit columns for readability
            sheet.AutoFitColumns();

            // Save the workbook (lifecycle rule)
            workbook.Save("ThemeAccentContrastMatrix.xlsx");
        }
    }
}
