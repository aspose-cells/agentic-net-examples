// Title: Generate a macro‑free Excel template with custom theme colors using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set custom theme colors for a new workbook | How to create a macro‑free Excel template and define accent colors with Aspose.Cells | Aspose.Cells Workbook.SetThemeColor example for saving a themed template file | Create an Excel template with predefined hyperlink colors using Aspose.Cells .NET | Generate an XLSX template workbook without macros and custom theme palette in C#
// Tags: Workbook.SetThemeColor custom palette | macro‑free template workbook Aspose.Cells | custom accent theme colors .NET | hyperlink theme color configuration Aspose.Cells | save themed workbook as Xlsx

using System;
using System.Drawing;
using Aspose.Cells;

namespace TemplateGenerator
{
    // The sample creates a new macro‑free Workbook, applies custom RGB values to the six accent theme colors and to hyperlink colors via Workbook.SetThemeColor, then saves the result as 'Template.xlsx' in XLSX format, handling any exceptions that may arise.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (macro‑free by default)
                Workbook workbook = new Workbook();

                // Define custom theme colors using the Workbook.SetThemeColor method
                // Adjust the RGB values as needed
                workbook.SetThemeColor(ThemeColorType.Accent1, Color.FromArgb(0, 112, 192));   // Blue
                workbook.SetThemeColor(ThemeColorType.Accent2, Color.FromArgb(255, 192, 0));   // Orange
                workbook.SetThemeColor(ThemeColorType.Accent3, Color.FromArgb(112, 173, 71)); // Green
                workbook.SetThemeColor(ThemeColorType.Accent4, Color.FromArgb(255, 0, 0));     // Red
                workbook.SetThemeColor(ThemeColorType.Accent5, Color.FromArgb(128, 0, 128));   // Purple
                workbook.SetThemeColor(ThemeColorType.Accent6, Color.FromArgb(0, 176, 80));    // Dark Green

                // Optionally, set hyperlink and followed hyperlink colors
                workbook.SetThemeColor(ThemeColorType.Hyperlink, Color.FromArgb(0, 0, 255));          // Standard blue
                workbook.SetThemeColor(ThemeColorType.FollowedHyperlink, Color.FromArgb(128, 0, 128)); // Purple

                // Define output file path
                string outputPath = "Template.xlsx";

                // Save the template workbook to a file
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
