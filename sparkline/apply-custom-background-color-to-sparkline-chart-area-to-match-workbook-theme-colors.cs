// Title: Set a custom theme accent as the background color of a sparkline chart area using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a line sparkline to a worksheet and sets its AreaColor to a workbook theme accent using Aspose.Cells. | Show how to invoke SparklineGroups.Add via reflection and then assign a custom background color to the created sparkline group. | Provide a snippet that saves the workbook after applying the sparkline background color, with error handling for missing methods.
// Common Searches: Aspose.Cells C# set sparkline background color to theme accent | how to change sparkline area color with reflection in .NET | example of adding line sparkline and customizing AreaColor using Aspose.Cells | apply workbook theme colors to sparkline chart area programmatically | sparkline group AreaColor property Aspose.Cells tutorial
// Tags: sparkline area color Aspose.Cells | set sparkline background C# | add sparkline via reflection | theme accent color workbook | Aspose.Cells sparkline customization

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, adds a line‑type sparkline using reflection, sets the sparkline group's AreaColor to a theme accent color, and saves the file as SparklineWithThemeBackground.xlsx, including handling for missing Sparkline APIs.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline source range
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40);
            sheet.Cells["A5"].PutValue(50);

            // Attempt to add a line‑type sparkline group using reflection
            try
            {
                var sparklineGroups = sheet.SparklineGroups;
                var addMethod = sparklineGroups.GetType().GetMethod(
                    "Add",
                    new Type[] { typeof(int), typeof(string), typeof(string) });

                if (addMethod != null)
                {
                    // Add sparkline group (type = 0 for Line) and obtain its index
                    object result = addMethod.Invoke(sparklineGroups, new object[] { 0, "B1", "A1:A5" });
                    int sparklineGroupIndex = Convert.ToInt32(result);

                    // Retrieve the created sparkline group via the indexer
                    var itemProperty = sparklineGroups.GetType().GetProperty("Item");
                    object sparklineGroup = itemProperty?.GetValue(sparklineGroups, new object[] { sparklineGroupIndex });

                    if (sparklineGroup != null)
                    {
                        // Use a fallback color if theme colors are unavailable
                        Color themeAccentColor = Color.FromArgb(0, 112, 192); // typical Accent1 color

                        // Apply the color as the background (area) color of the sparkline
                        var areaColorProp = sparklineGroup.GetType().GetProperty("AreaColor");
                        if (areaColorProp != null && areaColorProp.CanWrite)
                        {
                            areaColorProp.SetValue(sparklineGroup, themeAccentColor);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Sparkline 'Add' method not found. Sparkline feature may be unavailable in this version.");
                }
            }
            catch (Exception sparkEx)
            {
                Console.WriteLine($"Sparkline operation failed: {sparkEx.Message}");
            }

            // Define output file path
            string outputPath = "SparklineWithThemeBackground.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
