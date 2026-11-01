// Title: Replace Dark1 theme color with a custom hexadecimal value and save the workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change Dark1 theme color to a hex value | set custom hex color for Excel theme Dark1 programmatically | modify workbook theme color using reflection in Aspose.Cells | save Excel file after updating Dark1 theme color with Aspose.Cells | convert hex string to System.Drawing.Color for Aspose.Cells theme
// Tags: Aspose.Cells set Dark1 theme color | Excel theme color update C# | hex to Color conversion Aspose.Cells | save workbook after theme modification Aspose.Cells | reflection based Theme.SetColor method

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsThemeExample
{
    // The example loads an existing workbook, uses reflection to locate the ThemeColorType enum and the Theme object's SetColor method, converts a user‑provided hex string to a System.Drawing.Color, replaces the Dark1 theme color, ensures the output directory exists, and saves the modified workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // User‑specified hexadecimal color for Dark1 (e.g., "#FF5733")
            string dark1Hex = "#FF5733";

            // Paths for the source and destination workbooks
            string sourcePath = "input.xlsx";
            string destinationPath = "output.xlsx";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(sourcePath);

                // Attempt to modify the theme color using reflection (compatible with multiple Aspose.Cells versions)
                try
                {
                    var theme = workbook.Theme;
                    if (theme != null)
                    {
                        // Resolve the ThemeColorType enum type from the assembly that defines the theme
                        Type enumType = theme.GetType().Assembly.GetType("Aspose.Cells.ThemeColorType");
                        if (enumType != null)
                        {
                            // Parse the enum value for Dark1
                            object dark1Enum = Enum.Parse(enumType, "Dark1", ignoreCase: true);

                            // Locate the SetColor method that accepts (ThemeColorType, Color)
                            var setColorMethod = theme.GetType().GetMethod(
                                "SetColor",
                                new Type[] { enumType, typeof(Color) });

                            if (setColorMethod != null)
                            {
                                // Convert the hex string to a Color object
                                Color newColor = ColorTranslator.FromHtml(dark1Hex);

                                // Invoke SetColor to replace the Dark1 theme color
                                setColorMethod.Invoke(theme, new object[] { dark1Enum, newColor });
                            }
                            else
                            {
                                Console.WriteLine("SetColor method not found on Theme object.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("ThemeColorType enum not found. Theme modification skipped.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Workbook does not contain a Theme. Theme modification skipped.");
                    }
                }
                catch (Exception ex)
                {
                    // Handle any errors that occur while attempting to modify the theme
                    Console.WriteLine($"Error modifying theme: {ex.Message}");
                }

                // Ensure the destination directory exists
                string destDir = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                // Save the modified workbook
                workbook.Save(destinationPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {destinationPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
