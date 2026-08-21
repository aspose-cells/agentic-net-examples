// Title: Generate CSS from Excel Theme Accent Colors with Aspose.Cells for .NET (C#)
// Description: C# sample that loads or creates an Excel workbook, reads the six theme accent colors via Aspose.Cells Workbook.GetThemeColor, converts each System.Drawing.Color to a hex string, and writes CSS classes for text and background colors to a stylesheet. The workbook can then be saved for further processing.
// Keywords: Aspose.Cells | C# | .NET | Excel theme colors | ThemeColorType | GetThemeColor | CSS generation | hex color conversion | web styling from Excel | automated stylesheet
// Common Searches: Aspose.Cells extract theme accent colors C# | convert Excel theme colors to CSS | generate stylesheet from workbook theme | C# code to get Excel theme colors as hex | create CSS classes from Excel theme using Aspose
// Developer Intent: Read the workbook’s theme accent colors and output a CSS file with matching text‑color and background‑color classes.
// Use Cases: Synchronize web page styling with Excel report themes for consistent branding. | Automate CSS creation for multiple workbooks in a reporting pipeline. | Include generated CSS in CI/CD pipelines alongside published Excel files for front‑end developers.
// AI Prompts: Write a C# method that takes a Workbook and returns a dictionary of ThemeColorType to hex color strings using Aspose.Cells. | Provide code to extract all theme colors (accent, hyperlink, text1, text2) and generate an SCSS file with variables. | Explain how to detect a workbook without a custom theme and fall back to default theme colors when creating CSS.

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;
using System.Text;

// C# sample that loads or creates an Excel workbook, reads the six theme accent colors via Aspose.Cells Workbook.GetThemeColor, converts each System.Drawing.Color to a hex string, and writes CSS classes for text and background colors to a stylesheet. The workbook can then be saved for further processing.
class ExtractThemeColorsToCss
{
    static void Main()
    {
        // Create a new workbook (or load an existing one with a theme)
        Workbook workbook = new Workbook();
        // Example of loading a workbook with a custom theme:
        // Workbook workbook = new Workbook("input.xlsx");

        // Define accent names and corresponding ThemeColorType values
        string[] accentNames = { "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6" };
        ThemeColorType[] accentTypes = {
            ThemeColorType.Accent1,
            ThemeColorType.Accent2,
            ThemeColorType.Accent3,
            ThemeColorType.Accent4,
            ThemeColorType.Accent5,
            ThemeColorType.Accent6
        };

        // Build CSS content with classes for text color and background color
        StringBuilder cssBuilder = new StringBuilder();
        cssBuilder.AppendLine("/* Generated CSS for workbook theme accent colors */");

        for (int i = 0; i < accentTypes.Length; i++)
        {
            // Retrieve the accent color from the workbook theme
            Color accentColor = workbook.GetThemeColor(accentTypes[i]);

            // Convert the Color to a hex string (e.g., #FF00FF)
            string hexColor = ColorTranslator.ToHtml(accentColor);

            // Create CSS classes
            string className = accentNames[i].ToLower(); // e.g., accent1
            cssBuilder.AppendLine($".{className} {{ color: {hexColor}; }}");
            cssBuilder.AppendLine($".bg-{className} {{ background-color: {hexColor}; }}");
        }

        // Write the CSS to a file
        string cssFilePath = "theme-accent-colors.css";
        File.WriteAllText(cssFilePath, cssBuilder.ToString());

        // Save the workbook if further processing is required
        workbook.Save("output.xlsx");
    }
}
