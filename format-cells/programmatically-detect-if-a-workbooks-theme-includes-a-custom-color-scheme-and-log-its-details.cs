// Title: Programmatically detect a custom theme color scheme in an Excel workbook and log each theme color using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells check if workbook theme has custom color scheme | How to read and list theme colors from an .xlsx file with Aspose.Cells | Detect custom Excel theme programmatically using Aspose.Cells .NET | Retrieve ThemeColorType values from a workbook's theme in C# | Log custom theme colors from an Excel file using Aspose.Cells reflection
// Tags: custom theme detection Aspose.Cells .NET | retrieve workbook theme colors C# | Aspose.Cells ThemeColorType enumeration | reflection access Aspose.Cells Theme property | log Excel theme color values Aspose

using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, uses reflection to obtain the Theme and its ColorScheme, checks the IsCustom flag, and when a custom scheme is present iterates over all ThemeColorType entries, outputting each color's RGB hex value. It also includes file‑existence checks and graceful error handling.
class ThemeDetection
{
    static void Main()
    {
        // Path to the workbook to be examined
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Attempt to obtain the Theme object via reflection (compatible with versions that may not expose Theme directly)
            PropertyInfo themeProp = workbook.GetType().GetProperty("Theme", BindingFlags.Public | BindingFlags.Instance);
            if (themeProp == null)
            {
                Console.WriteLine("The workbook does not expose a Theme property in this Aspose.Cells version.");
                return;
            }

            object themeObj = themeProp.GetValue(workbook);
            if (themeObj == null)
            {
                Console.WriteLine("The workbook does not contain a theme.");
                return;
            }

            // Obtain the ColorScheme from the Theme object
            PropertyInfo colorSchemeProp = themeObj.GetType().GetProperty("ColorScheme", BindingFlags.Public | BindingFlags.Instance);
            if (colorSchemeProp == null)
            {
                Console.WriteLine("Unable to retrieve the theme color scheme.");
                return;
            }

            object colorSchemeObj = colorSchemeProp.GetValue(themeObj);
            if (colorSchemeObj == null)
            {
                Console.WriteLine("The theme color scheme is null.");
                return;
            }

            // Check if the color scheme is custom
            PropertyInfo isCustomProp = colorSchemeObj.GetType().GetProperty("IsCustom", BindingFlags.Public | BindingFlags.Instance);
            bool isCustom = isCustomProp != null && (bool)(isCustomProp.GetValue(colorSchemeObj) ?? false);

            if (isCustom)
            {
                Console.WriteLine("The workbook contains a custom color scheme.");

                // Get the GetColor method to retrieve colors for each ThemeColorType
                MethodInfo getColorMethod = colorSchemeObj.GetType().GetMethod("GetColor", new[] { typeof(ThemeColorType) });
                if (getColorMethod == null)
                {
                    Console.WriteLine("Unable to retrieve colors from the color scheme.");
                    return;
                }

                // Iterate through all possible theme color types and log their RGB values
                foreach (ThemeColorType colorType in Enum.GetValues(typeof(ThemeColorType)))
                {
                    object result = getColorMethod.Invoke(colorSchemeObj, new object[] { colorType });
                    if (result is Color color)
                    {
                        Console.WriteLine($"{colorType}: #{color.R:X2}{color.G:X2}{color.B:X2}");
                    }
                }
            }
            else
            {
                Console.WriteLine("The workbook uses the default theme color scheme.");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
