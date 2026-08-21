// Title: C# – Get Hyperlink Theme Color RGB from an Excel Workbook with Aspose.Cells
// Description: Loads an Excel file, calls Workbook.GetThemeColor(ThemeColorType.Hyperlink) to obtain the System.Drawing.Color for the Hyperlink theme, and reads its ARGB values for display or further processing.
// Keywords: Aspose.Cells | C# | .NET | Workbook.GetThemeColor | Hyperlink theme color | RGB | Excel theme colors | retrieve theme color | Color object | Excel hyperlink color
// Common Searches: Aspose.Cells get hyperlink theme color C# | How to read Excel hyperlink theme RGB using Aspose | Workbook.GetThemeColor Hyperlink example | C# retrieve Excel theme color for hyperlinks | Get ARGB of hyperlink theme in Aspose.Cells
// Developer Intent: Extract the RGB (and alpha) components of the Hyperlink theme color from a loaded Excel workbook using Aspose.Cells for .NET.
// Use Cases: Show the workbook's hyperlink color in a custom UI to match the document's styling. | Validate that the hyperlink color complies with corporate branding before applying conditional formatting. | Log ARGB values of the hyperlink theme across multiple files for audit or migration purposes.
// AI Prompts: Write C# code with Aspose.Cells that retrieves the Hyperlink theme color and returns it as a hexadecimal string. | Demonstrate how to change the Hyperlink theme color to a custom RGB value in an existing workbook using Aspose.Cells. | Provide robust error‑handling patterns when calling Workbook.GetThemeColor for unsupported or missing theme entries.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel file, calls Workbook.GetThemeColor(ThemeColorType.Hyperlink) to obtain the System.Drawing.Color for the Hyperlink theme, and reads its ARGB values for display or further processing.
    public class RetrieveHyperlinkThemeColor
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string filePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(filePath);

                // Get the theme color for Hyperlink
                Color hyperlinkColor = workbook.GetThemeColor(ThemeColorType.Hyperlink);

                // Output the ARGB components of the theme color
                Console.WriteLine($"Hyperlink Theme Color - A:{hyperlinkColor.A}, R:{hyperlinkColor.R}, G:{hyperlinkColor.G}, B:{hyperlinkColor.B}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
