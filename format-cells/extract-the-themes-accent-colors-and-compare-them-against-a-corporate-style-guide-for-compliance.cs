using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeComplianceCheck
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default theme will be loaded)
            Workbook workbook = new Workbook();

            // Extract the six accent colors from the workbook's theme
            Color[] workbookAccentColors = new Color[6];
            for (int i = 0; i < 6; i++)
            {
                // ThemeColorType.Accent1 has the enum value 4, so we offset by i
                ThemeColorType accentType = (ThemeColorType)((int)ThemeColorType.Accent1 + i);
                workbookAccentColors[i] = workbook.GetThemeColor(accentType);
            }

            // Define the corporate style guide accent colors (example values)
            Color[] corporateAccentColors = new Color[6]
            {
                Color.FromArgb(255, 0, 112, 192),   // Corporate Accent1
                Color.FromArgb(255, 255, 192, 0),   // Corporate Accent2
                Color.FromArgb(255, 112, 173, 71),  // Corporate Accent3
                Color.FromArgb(255, 237, 125, 49),  // Corporate Accent4
                Color.FromArgb(255, 191, 0, 0),     // Corporate Accent5
                Color.FromArgb(255, 112, 48, 160)   // Corporate Accent6
            };

            // Compare each workbook accent color with the corporate guide
            Console.WriteLine("Theme Accent Color Compliance Report:");
            for (int i = 0; i < 6; i++)
            {
                bool isMatch = workbookAccentColors[i].ToArgb() == corporateAccentColors[i].ToArgb();
                string status = isMatch ? "Compliant" : "Non‑compliant";
                Console.WriteLine($"Accent{i + 1}: Workbook = {workbookAccentColors[i]}, " +
                                  $"Corporate = {corporateAccentColors[i]} => {status}");
            }

            // Optionally save the workbook (demonstrates lifecycle rule usage)
            workbook.Save("ThemeComplianceResult.xlsx");
        }
    }
}