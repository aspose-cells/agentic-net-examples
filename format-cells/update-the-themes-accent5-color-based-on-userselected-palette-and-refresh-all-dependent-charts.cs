// Title: Update the workbook's Accent5 theme color and refresh dependent charts with Aspose.Cells for .NET
// AI Prompts: Set the Accent5 theme color of an Excel workbook to a specific RGB value using Aspose.Cells in C#. | Load an existing .xlsx file (or create a new workbook if it does not exist), apply a custom Accent5 color, and save the workbook. | Catch and log errors when applying a theme color and verify the output folder exists before saving.
// Common Searches: Aspose.Cells change Accent5 theme color programmatically C# | How to refresh charts after updating Excel theme colors with Aspose.Cells | SetThemeColor method example for custom RGB in .NET workbook | Create workbook if file not found and apply theme color using Aspose.Cells | Save Excel file after modifying theme colors and ensure output folder exists
// Tags: Aspose.Cells SetThemeColor Accent5 | update Excel theme color C# | refresh charts after theme change Aspose.Cells | create workbook if missing Aspose.Cells | ensure output directory before saving workbook

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook (or creates a new one if the file is missing), sets the theme's Accent5 color to a user‑specified RGB value, ensures the output directory exists, and saves the workbook while handling potential exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load existing workbook or create a new one if the file is missing
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // User‑selected Accent5 color (example: red)
                Color userAccent5 = Color.FromArgb(255, 0, 0);

                // Update the theme's Accent5 color using the supported API
                try
                {
                    workbook.SetThemeColor(ThemeColorType.Accent5, userAccent5);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to set theme color. {ex.Message}");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
