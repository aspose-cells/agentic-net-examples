// Title: Convert HTML to Excel with Aspose.Cells and Apply CSS Variable Theme (C#)
// Description: C# sample that reads an HTML file, extracts CSS custom properties (e.g., --primary-color, --secondary-color) with a regex, loads the HTML into an Aspose.Cells Workbook, applies the extracted colors to the workbook's default style, and saves the result as an XLSX file.
// Keywords: Aspose.Cells HTML to XLSX | C# convert HTML to Excel | apply CSS variables to Excel workbook | custom theme Aspose.Cells | extract CSS custom properties C# | load HTML with LoadOptions Aspose | default style colors Excel | brand‑consistent Excel export
// Common Searches: How to load HTML into Aspose.Cells and keep CSS colors | C# extract CSS variables from an HTML file for Excel styling | Apply a custom theme to a workbook created from HTML using Aspose.Cells | Set background and font colors in Aspose.Cells based on CSS variables | Convert HTML template to branded Excel file C#
// Developer Intent: Load an HTML document into a Workbook, read its CSS custom properties, and style the workbook before saving as XLSX.
// Use Cases: Generate a brand‑consistent report by converting an HTML template with CSS variables into an Excel file that mirrors the original colors. | Automate data export where the visual theme is driven by CSS variables defined in the source HTML. | Transform marketing email HTML into an Excel sheet that reflects the email’s color scheme for analytics.
// AI Prompts: Write C# code using Aspose.Cells to load an HTML file, extract CSS variables via regex, and apply those colors to the workbook's default style. | Show how to safely handle invalid CSS color values when applying them to an Aspose.Cells workbook. | Extend the example to map additional CSS variables such as --accent-color to cell borders or header styles.

using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

// C# sample that reads an HTML file, extracts CSS custom properties (e.g., --primary-color, --secondary-color) with a regex, loads the HTML into an Aspose.Cells Workbook, applies the extracted colors to the workbook's default style, and saves the result as an XLSX file.
class HtmlToExcelWithCustomTheme
{
    static void Main()
    {
        // Paths for the source HTML and the target Excel file
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        // -----------------------------------------------------------------
        // 1. Read the HTML content and extract CSS custom properties (variables)
        // -----------------------------------------------------------------
        string htmlContent = File.ReadAllText(htmlPath);

        // Regex to capture CSS variables defined like: --primary-color: #ff0000;
        Regex cssVarRegex = new Regex(@"--(?<name>[\w-]+)\s*:\s*(?<value>[^;]+);");
        MatchCollection matches = cssVarRegex.Matches(htmlContent);

        // Store the extracted variables in a dictionary for easy lookup
        Dictionary<string, string> cssVariables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match match in matches)
        {
            string name = match.Groups["name"].Value.Trim();
            string value = match.Groups["value"].Value.Trim();
            cssVariables[name] = value;
        }

        // -----------------------------------------------------------------
        // 2. Load the HTML file into an Aspose.Cells Workbook
        // -----------------------------------------------------------------
        // Use LoadOptions with LoadFormat.Html to correctly interpret the HTML.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // -----------------------------------------------------------------
        // 3. Apply a simple custom theme based on the extracted CSS variables
        // -----------------------------------------------------------------
        // Example convention:
        //   --primary-color   -> background color for the default style
        //   --secondary-color -> font color for the default style
        // Adjust as needed for more complex theming.

        // Apply background color if defined
        if (cssVariables.TryGetValue("primary-color", out string primaryColorValue))
        {
            try
            {
                Color bgColor = ColorTranslator.FromHtml(primaryColorValue);
                Style defaultStyle = workbook.DefaultStyle;
                defaultStyle.ForegroundColor = bgColor;
                defaultStyle.Pattern = BackgroundType.Solid;
                workbook.DefaultStyle = defaultStyle;
            }
            catch
            {
                // Invalid color format – ignore and continue
            }
        }

        // Apply font color if defined
        if (cssVariables.TryGetValue("secondary-color", out string secondaryColorValue))
        {
            try
            {
                Color fontColor = ColorTranslator.FromHtml(secondaryColorValue);
                Style defaultStyle = workbook.DefaultStyle;
                defaultStyle.Font.Color = fontColor;
                workbook.DefaultStyle = defaultStyle;
            }
            catch
            {
                // Invalid color format – ignore and continue
            }
        }

        // -----------------------------------------------------------------
        // 4. Save the workbook as an Excel file (XLSX)
        // -----------------------------------------------------------------
        workbook.Save(excelPath, SaveFormat.Xlsx);

        Console.WriteLine($"HTML has been converted to Excel and saved to '{excelPath}'.");
    }
}
