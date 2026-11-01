// Title: Apply a custom corporate color theme to an existing Excel workbook with Aspose.Cells for .NET
// Common Searches: how to change Excel theme accent colors using Aspose.Cells in C# | set corporate brand palette for workbook theme with Aspose.Cells .NET | Aspose.Cells SetThemeColor example for customizing Excel themes | apply custom theme to existing .xlsx file using Aspose.Cells | C# code to modify Excel theme colors programmatically
// Tags: Aspose.Cells SetThemeColor API | custom Excel theme colors .NET | programmatic workbook theme customization | apply corporate palette to Excel workbook | theme accent color modification using Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsThemeExample
{
    // The sample loads an existing workbook, defines corporate brand colors, assigns them to the workbook's theme accents (Accent1‑Accent6, Hyperlink, FollowedHyperlink) via Workbook.SetThemeColor, and saves the file with the new custom theme applied.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "ReportInput.xlsx";
                const string outputPath = "ReportOutput.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Define corporate brand colors
                Color brandPrimary   = ColorTranslator.FromHtml("#003366"); // Dark Blue
                Color brandSecondary = ColorTranslator.FromHtml("#FFCC00"); // Gold
                Color brandAccent1   = ColorTranslator.FromHtml("#0099CC"); // Light Blue
                Color brandAccent2   = ColorTranslator.FromHtml("#66CC66"); // Green
                Color brandAccent3   = ColorTranslator.FromHtml("#CC3333"); // Red
                Color brandAccent4   = ColorTranslator.FromHtml("#9966CC"); // Purple
                Color brandAccent5   = ColorTranslator.FromHtml("#FF6600"); // Orange
                Color brandAccent6   = ColorTranslator.FromHtml("#CCCCCC"); // Light Gray

                // Apply corporate colors to the workbook theme accents
                workbook.SetThemeColor(ThemeColorType.Accent1, brandPrimary);
                workbook.SetThemeColor(ThemeColorType.Accent2, brandSecondary);
                workbook.SetThemeColor(ThemeColorType.Accent3, brandAccent1);
                workbook.SetThemeColor(ThemeColorType.Accent4, brandAccent2);
                workbook.SetThemeColor(ThemeColorType.Accent5, brandAccent3);
                workbook.SetThemeColor(ThemeColorType.Accent6, brandAccent4);

                // Optionally set hyperlink colors to match brand palette
                workbook.SetThemeColor(ThemeColorType.Hyperlink, brandAccent5);
                workbook.SetThemeColor(ThemeColorType.FollowedHyperlink, brandAccent6);

                // Save the workbook with the custom theme applied
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log or display the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
