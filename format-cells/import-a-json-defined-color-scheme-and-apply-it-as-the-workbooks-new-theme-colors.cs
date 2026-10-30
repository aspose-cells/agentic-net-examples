// Title: Load a JSON color palette and set it as the workbook theme using Aspose.Cells for .NET
// Common Searches: how to load theme colors from a JSON string and apply them to an Excel workbook with Aspose.Cells C# | Aspose.Cells set custom theme colors from a dictionary of hex values | C# example using SetThemeColor and ThemeColorType to import a JSON color scheme | apply JSON-defined hex colors to a workbook theme in Aspose.Cells .NET | convert JSON color map to System.Drawing.Color for Excel theme with Aspose.Cells
// Tags: json color palette to workbook theme Aspose.Cells | setthemecolor method usage C# | themecolortype mapping from json Aspose.Cells | hex to System.Drawing.Color conversion Aspose.Cells | custom Excel theme creation Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.Json;

// This example shows how to deserialize a JSON string containing hex values for theme colors into a dictionary, convert each hex value to a System.Drawing.Color, map the color name to the ThemeColorType enum, and apply the colors to a new Workbook using the SetThemeColor API of Aspose.Cells for .NET, then save the themed workbook.
class Program
{
    static void Main()
    {
        try
        {
            // JSON string that defines the theme colors
            string json = @"{
                ""Accent1"": ""#FF5733"",
                ""Accent2"": ""#33FF57"",
                ""Accent3"": ""#3357FF"",
                ""Accent4"": ""#FF33A1"",
                ""Accent5"": ""#A133FF"",
                ""Accent6"": ""#33FFF5"",
                ""Hyperlink"": ""#0000FF"",
                ""FollowedHyperlink"": ""#800080""
            }";

            // Deserialize JSON into a dictionary of color name -> hex value
            var colorMap = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
            if (colorMap == null)
                throw new InvalidOperationException("Failed to deserialize theme colors.");

            // Create a new workbook
            var workbook = new Workbook();

            // Apply each color to the workbook's theme
            foreach (var kvp in colorMap)
            {
                // Convert the string key to the ThemeColorType enum
                if (Enum.TryParse<ThemeColorType>(kvp.Key, out var themeColor))
                {
                    // Convert hex string to a System.Drawing.Color
                    Color color = ColorTranslator.FromHtml(kvp.Value);
                    // Set the theme color using the Workbook API
                    workbook.SetThemeColor(themeColor, color);
                }
            }

            // Save the workbook with the new theme applied
            workbook.Save("ThemedWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
