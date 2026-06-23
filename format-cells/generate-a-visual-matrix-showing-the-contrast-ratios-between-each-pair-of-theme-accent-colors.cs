using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeAccentContrastMatrix
{
    class Program
    {
        // Compute relative luminance of a color according to WCAG
        static double GetLuminance(Color color)
        {
            // Convert sRGB components to linear values
            double R = color.R / 255.0;
            double G = color.G / 255.0;
            double B = color.B / 255.0;

            R = (R <= 0.03928) ? R / 12.92 : Math.Pow((R + 0.055) / 1.055, 2.4);
            G = (G <= 0.03928) ? G / 12.92 : Math.Pow((G + 0.055) / 1.055, 2.4);
            B = (B <= 0.03928) ? B / 12.92 : Math.Pow((B + 0.055) / 1.055, 2.4);

            // Relative luminance
            return 0.2126 * R + 0.7152 * G + 0.0722 * B;
        }

        // Compute contrast ratio between two colors
        static double GetContrastRatio(Color c1, Color c2)
        {
            double L1 = GetLuminance(c1);
            double L2 = GetLuminance(c2);

            // Ensure L1 is the lighter color
            if (L2 > L1)
            {
                double temp = L1;
                L1 = L2;
                L2 = temp;
            }

            return (L1 + 0.05) / (L2 + 0.05);
        }

        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the accent types we are interested in
            ThemeColorType[] accents = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            // Retrieve the actual colors for each accent
            Color[] accentColors = new Color[accents.Length];
            for (int i = 0; i < accents.Length; i++)
            {
                accentColors[i] = workbook.GetThemeColor(accents[i]);
            }

            // Write headers (first row and first column)
            for (int i = 0; i < accents.Length; i++)
            {
                // Header row (starting from column B)
                sheet.Cells[0, i + 1].PutValue(accents[i].ToString());
                // Header column (starting from row 2)
                sheet.Cells[i + 1, 0].PutValue(accents[i].ToString());
            }

            // Fill matrix with contrast ratios
            for (int row = 0; row < accents.Length; row++)
            {
                for (int col = 0; col < accents.Length; col++)
                {
                    double ratio = GetContrastRatio(accentColors[row], accentColors[col]);
                    // Round to two decimal places for readability
                    sheet.Cells[row + 1, col + 1].PutValue(Math.Round(ratio, 2));
                }
            }

            // Optional: format the matrix for better visual appearance
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;
            headerStyle.VerticalAlignment = TextAlignmentType.Center;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.ForegroundColor = Color.LightGray;

            // Apply header style
            for (int i = 0; i <= accents.Length; i++)
            {
                // Top header row
                sheet.Cells[0, i].SetStyle(headerStyle);
                // Left header column
                sheet.Cells[i, 0].SetStyle(headerStyle);
            }

            // Auto-fit columns for readability
            sheet.AutoFitColumns();

            // Save the workbook (lifecycle rule)
            workbook.Save("ThemeAccentContrastMatrix.xlsx");
        }
    }
}