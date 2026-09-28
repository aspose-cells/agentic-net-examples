// Title: How to lighten the Light2 background color of an Excel workbook theme and set dark text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, accesses the workbook's Theme ColorScheme via reflection, changes the Light2 color to a light gray, and saves the file. | Write a reusable C# method that safely updates any theme color (e.g., Light2, Text1) in an Aspose.Cells workbook, handling missing properties with reflection. | Provide a step‑by‑step guide to adjust the Excel theme's background shade and primary text color programmatically using Aspose.Cells for .NET.
// Common Searches: c# aspocells change Light2 theme color to light gray | set primary text (Text1) to black in Excel theme using Aspose.Cells | use reflection to modify workbook theme colors Aspose.Cells .NET | adjust Excel theme background shade programmatically with Aspose.Cells
// Tags: update Light2 color Aspose.Cells theme | set Text1 black Aspose.Cells workbook | reflection access Theme ColorScheme C# | modify Excel theme background shade .NET | load and save workbook Aspose.Cells C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The program loads or creates an Excel workbook, uses reflection to reach the Theme ColorScheme, sets Light2 to a very light gray (RGB 240,240,240) and Text1 to black, then saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Attempt to modify the workbook's theme color scheme (if supported)
            try
            {
                // Use reflection to avoid compile‑time dependency on ThemeColorScheme
                var themeObj = workbook.Theme;
                var colorSchemeProp = themeObj?.GetType().GetProperty("ColorScheme");
                if (colorSchemeProp != null)
                {
                    var colorScheme = colorSchemeProp.GetValue(themeObj);
                    var light2Prop = colorScheme?.GetType().GetProperty("Light2");
                    var text1Prop = colorScheme?.GetType().GetProperty("Text1");

                    if (light2Prop != null && text1Prop != null)
                    {
                        // Set Light2 to a very light gray
                        light2Prop.SetValue(colorScheme, Color.FromArgb(255, 240, 240, 240));
                        // Set Text1 (primary dark text) to black
                        text1Prop.SetValue(colorScheme, Color.FromArgb(255, 0, 0, 0));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Theme modification skipped: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
