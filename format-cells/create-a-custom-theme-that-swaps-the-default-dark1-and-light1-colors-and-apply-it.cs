// Title: Swap the default Dark1 and Light1 colors in an Excel workbook theme using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to obtain the workbook Theme via reflection, exchange the Dark1 and Light1 colors, and save the file as XLSX. | Demonstrate a safe approach to modify a workbook's Theme.ColorScheme with reflection and include a fallback when the Theme API is unavailable in Aspose.Cells .NET.
// Common Searches: how to change Dark1 color in an Excel theme with Aspose.Cells C# | using reflection to access Theme.ColorScheme in Aspose.Cells .NET | swap default theme colors Dark1 and Light1 when generating XLSX with Aspose.Cells | programmatic Excel theme customization fallback if Theme API is missing | C# example for exchanging theme colors in an Aspose.Cells workbook
// Tags: swap theme dark1 light1 Aspose.Cells | reflection based theme color change .NET | custom Excel theme color scheme Aspose.Cells | modify workbook theme colors C# | fallback handling missing Theme API Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;

// The example creates a new Workbook, uses reflection to access the Theme and its ColorScheme, swaps the Dark1 and Light1 colors, gracefully handles cases where the Theme API is not present, and saves the result as CustomThemeSwapped.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Attempt to modify theme colors via reflection (works only if Theme API is available)
            try
            {
                // Get the Theme property dynamically
                PropertyInfo themeProp = workbook.GetType().GetProperty("Theme", BindingFlags.Public | BindingFlags.Instance);
                object themeObj = themeProp?.GetValue(workbook);
                if (themeObj != null)
                {
                    // Get the ColorScheme property from the Theme object
                    PropertyInfo colorSchemeProp = themeObj.GetType().GetProperty("ColorScheme", BindingFlags.Public | BindingFlags.Instance);
                    object colorSchemeObj = colorSchemeProp?.GetValue(themeObj);
                    if (colorSchemeObj != null)
                    {
                        // Access Dark1 and Light1 color properties
                        PropertyInfo dark1Prop = colorSchemeObj.GetType().GetProperty("Dark1", BindingFlags.Public | BindingFlags.Instance);
                        PropertyInfo light1Prop = colorSchemeObj.GetType().GetProperty("Light1", BindingFlags.Public | BindingFlags.Instance);

                        if (dark1Prop != null && light1Prop != null)
                        {
                            // Store original colors
                            Color originalDark1 = (Color)dark1Prop.GetValue(colorSchemeObj);
                            Color originalLight1 = (Color)light1Prop.GetValue(colorSchemeObj);

                            // Swap the colors
                            dark1Prop.SetValue(colorSchemeObj, originalLight1);
                            light1Prop.SetValue(colorSchemeObj, originalDark1);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // If reflection fails (e.g., Theme API not present), continue without theme changes
                Console.WriteLine($"Theme manipulation skipped: {ex.Message}");
            }

            // Define output file path
            string outputPath = "CustomThemeSwapped.xlsx";

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
