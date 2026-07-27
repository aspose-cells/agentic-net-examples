// Title: Convert HTML to Excel and Apply a CSS‑Variable Theme with Aspose.Cells for .NET
// Description: Load an HTML file into an Aspose.Cells Workbook, extract CSS custom properties via regex, map those variables to the workbook's default style and header row, then save the styled sheet as XLSX.
// Keywords: Aspose.Cells HTML to Excel | C# convert HTML to XLSX | CSS variables Excel styling | apply theme from HTML to workbook | extract CSS custom properties C# | default style Aspose.Cells | header row formatting Aspose
// Common Searches: Aspose.Cells convert HTML to Excel with styling | map CSS variables to Excel cell format C# | apply custom theme from HTML to XLSX using Aspose | extract --font-color and --bg-color from HTML for Excel | C# example: HTML to Excel with header style
// Developer Intent: Transform an HTML document into a styled Excel workbook by using CSS variables defined in the HTML as the source of the Excel theme.
// Use Cases: Generate a branded report by converting an HTML template that uses CSS variables into a matching Excel file. | Preserve the visual theme of a web‑based dashboard when exporting data to Excel for offline analysis. | Automate batch processing of HTML invoices, applying each file's CSS‑based colors to the resulting Excel sheets.
// AI Prompts: Provide C# code that loads an HTML file with Aspose.Cells, extracts CSS custom properties, and applies them to the workbook's default and header styles. | Explain how to map CSS variable values (e.g., --font-color, --bg-color) to Excel cell formatting using Aspose.Cells for .NET. | Suggest ways to extend the example to support row‑alternating colors, custom fonts, and conditional formatting based on additional CSS variables.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Drawing;
using Aspose.Cells;

namespace HtmlToExcelWithTheme
{
    // Load an HTML file into an Aspose.Cells Workbook, extract CSS custom properties via regex, map those variables to the workbook's default style and header row, then save the styled sheet as XLSX.
    class Program
    {
        static void Main()
        {
            // Paths for the source HTML and the target Excel file
            string htmlFilePath = "input.html";
            string excelFilePath = "output.xlsx";

            // Load the HTML file into a workbook (Aspose.Cells detects the format automatically)
            Workbook workbook = new Workbook(htmlFilePath);

            // Read the entire HTML content to extract CSS custom properties (variables)
            string htmlContent = File.ReadAllText(htmlFilePath);

            // Regex to match CSS variables defined like: --variable-name: value;
            Regex cssVarRegex = new Regex(@"--(?<name>[\w-]+)\s*:\s*(?<value>[^;]+);", RegexOptions.IgnoreCase);
            MatchCollection matches = cssVarRegex.Matches(htmlContent);

            // Store extracted variables in a dictionary for easy lookup
            Dictionary<string, string> cssVariables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in matches)
            {
                string name = match.Groups["name"].Value.Trim();
                string value = match.Groups["value"].Value.Trim();
                cssVariables[name] = value;
            }

            // Example: Apply a simple theme based on known variable names
            // --font-color   -> default font color
            // --bg-color     -> default cell background color
            // --header-bg    -> background color for the first row (header)
            // --header-font  -> font color for the header row
            Style defaultStyle = workbook.DefaultStyle;

            if (cssVariables.TryGetValue("font-color", out string fontColor))
            {
                defaultStyle.Font.Color = ColorTranslator.FromHtml(fontColor);
            }

            if (cssVariables.TryGetValue("bg-color", out string bgColor))
            {
                defaultStyle.ForegroundColor = ColorTranslator.FromHtml(bgColor);
                defaultStyle.Pattern = BackgroundType.Solid;
            }

            // Apply the modified default style back to the workbook
            workbook.DefaultStyle = defaultStyle;

            // Apply header styling if variables are present
            Worksheet sheet = workbook.Worksheets[0];
            Row headerRow = sheet.Cells.Rows[0];
            Style headerStyle = workbook.CreateStyle();

            if (cssVariables.TryGetValue("header-bg", out string headerBg))
            {
                headerStyle.ForegroundColor = ColorTranslator.FromHtml(headerBg);
                headerStyle.Pattern = BackgroundType.Solid;
            }

            if (cssVariables.TryGetValue("header-font", out string headerFont))
            {
                headerStyle.Font.Color = ColorTranslator.FromHtml(headerFont);
                headerStyle.Font.IsBold = true;
            }

            // Apply the header style to the entire first row
            foreach (Cell cell in headerRow)
            {
                cell.SetStyle(headerStyle);
            }

            // Save the workbook as an Excel file
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML has been converted to Excel and saved to '{excelFilePath}'.");
        }
    }
}
