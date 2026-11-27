// Title: Apply a corporate brand color theme to an existing Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: asp.net apply custom corporate theme to Excel workbook with Aspose.Cells C# | set ThemeColorType accent colors in an existing .xlsx using Aspose.Cells | programmatically change Excel theme colors in C# Aspose.Cells | load workbook and modify its color scheme with Aspose.Cells .NET | handle missing Theme API when customizing Excel theme in Aspose.Cells
// Tags: custom Excel theme Aspose.Cells C# | ThemeColorType SetColor usage | corporate brand colors workbook | dynamic theme manipulation Aspose.Cells | fallback when theme API unavailable

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // For ThemeColorType enum

namespace CorporateReportTheme
{
    // The example loads an existing or new workbook, accesses its Theme and ColorScheme via dynamic objects, defines corporate brand colors, assigns them to accent, hyperlink, background, and text slots using ThemeColorType.SetColor, and saves the file. If the Theme API is not present in the used Aspose.Cells version, the code logs the issue and continues without applying the custom theme.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputFile = "ReportInput.xlsx";
                const string outputFile = "ReportOutput_WithCustomTheme.xlsx";

                // Load the existing workbook if it exists; otherwise create a new one.
                Workbook workbook = File.Exists(inputFile) ? new Workbook(inputFile) : new Workbook();

                // Attempt to access the workbook's theme and its color scheme.
                // Use dynamic to avoid compile‑time dependency on Theme classes that may not be present in older versions.
                try
                {
                    dynamic theme = workbook.Theme;               // May throw if Theme property is unavailable
                    dynamic colorScheme = theme.ColorScheme;      // May throw if ColorScheme property is unavailable

                    // Define corporate brand colors.
                    Color primaryColor   = Color.FromArgb(0, 70, 127);   // Dark blue
                    Color secondaryColor = Color.FromArgb(255, 140, 0); // Orange
                    Color accent1Color   = Color.FromArgb(0, 120, 215); // Light blue
                    Color accent2Color   = Color.FromArgb(255, 215, 0); // Gold
                    Color accent3Color   = Color.FromArgb(0, 150, 0);   // Green
                    Color accent4Color   = Color.FromArgb(150, 0, 150); // Purple
                    Color accent5Color   = Color.FromArgb(200, 200, 200); // Light gray
                    Color accent6Color   = Color.FromArgb(50, 50, 50);   // Dark gray

                    // Apply colors to the theme's color scheme.
                    colorScheme.SetColor(ThemeColorType.Accent1, accent1Color);
                    colorScheme.SetColor(ThemeColorType.Accent2, accent2Color);
                    colorScheme.SetColor(ThemeColorType.Accent3, accent3Color);
                    colorScheme.SetColor(ThemeColorType.Accent4, accent4Color);
                    colorScheme.SetColor(ThemeColorType.Accent5, accent5Color);
                    colorScheme.SetColor(ThemeColorType.Accent6, accent6Color);
                    colorScheme.SetColor(ThemeColorType.Hyperlink, secondaryColor);
                    colorScheme.SetColor(ThemeColorType.FollowedHyperlink, primaryColor);

                    // Optionally set background and text colors.
                    colorScheme.SetColor(ThemeColorType.Background1, Color.White);
                    colorScheme.SetColor(ThemeColorType.Text1, Color.Black);
                    colorScheme.SetColor(ThemeColorType.Background2, Color.FromArgb(240, 240, 240));
                    colorScheme.SetColor(ThemeColorType.Text2, Color.FromArgb(80, 80, 80));
                }
                catch (Exception themeEx)
                {
                    // Theme manipulation is not supported in this version; continue without applying a custom theme.
                    Console.WriteLine($"Theme customization skipped: {themeEx.Message}");
                }

                // Save the workbook with the (potentially) applied theme.
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
