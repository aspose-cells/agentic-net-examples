// Title: Convert HTML tables with CSS background‑color classes to Excel while preserving cell colors using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads an HTML file, extracts CSS class background‑color definitions, loads the HTML into an Aspose.Cells Workbook, and applies those colors as solid fills to the matching worksheet cells. | Extend the sample to also detect inline style background‑color attributes on <td> or <th> elements and map them to Excel cell formatting. | Implement a reusable method that takes a Dictionary<string,string> of CSS class names to hex colors and a Worksheet, then applies the colors to cells based on each cell's class attribute.
// Common Searches: aspocells c# convert html table to xlsx preserving css background colors | how to map css class background-color to excel cell style using Aspose.Cells | load html with css styling into workbook and keep cell colors in .NET | extract css class colors from html and apply to excel cells programmatically
// Tags: html to xlsx conversion with css background colors Aspose.Cells | apply css class colors to worksheet cells C# | HtmlLoadOptions preserve styling Aspose.Cells | regex parse css background-color for excel formatting | set cell style foreground color from hex Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program reads 'input.html', extracts CSS class background‑color rules via regex, loads the HTML into an Aspose.Cells Workbook using HtmlLoadOptions, iterates over table rows and cells, applies the extracted colors as solid foreground fills to the corresponding worksheet cells, and saves the result as 'output.xlsx'.
class HtmlToExcelWithConditionalFormatting
{
    static void Main()
    {
        // Paths for input HTML and output Excel files
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        try
        {
            // -----------------------------------------------------------------
            // 1. Verify input HTML file exists and load its content into a string
            // -----------------------------------------------------------------
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Input HTML file not found: {htmlPath}");

            string htmlContent = File.ReadAllText(htmlPath);

            // -----------------------------------------------------------------
            // 2. Load the HTML into an Aspose.Cells Workbook
            //    HtmlLoadOptions preserves most of the CSS styling.
            // -----------------------------------------------------------------
            HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
            string tempHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".html");
            File.WriteAllText(tempHtmlPath, htmlContent);

            Workbook workbook;
            try
            {
                workbook = new Workbook(tempHtmlPath, loadOptions);
            }
            finally
            {
                // Clean up the temporary HTML file
                if (File.Exists(tempHtmlPath))
                {
                    try { File.Delete(tempHtmlPath); } catch { /* ignore cleanup errors */ }
                }
            }

            // -----------------------------------------------------------------
            // 3. Parse CSS classes that define background colors.
            //    Pattern extracts the class name and the hex color value.
            // -----------------------------------------------------------------
            var cssClassToColor = new Dictionary<string, string>();
            var cssMatches = Regex.Matches(
                htmlContent,
                @"\.([a-zA-Z0-9_-]+)\s*\{\s*background-color\s*:\s*(#[0-9A-Fa-f]{6})\s*;?\s*\}",
                RegexOptions.Multiline);

            foreach (Match match in cssMatches)
            {
                string className = match.Groups[1].Value;
                string colorHex = match.Groups[2].Value;
                cssClassToColor[className] = colorHex;
            }

            // -----------------------------------------------------------------
            // 4. Walk through the HTML table rows/cells using regex (no external HTML parser).
            //    Apply the corresponding background color to each cell.
            // -----------------------------------------------------------------
            var worksheet = workbook.Worksheets[0];
            int startRow = 0; // zero‑based index for the first row of the table

            var rowMatches = Regex.Matches(
                htmlContent,
                @"<tr[^>]*>(.*?)</tr>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            for (int r = 0; r < rowMatches.Count; r++)
            {
                string rowInnerHtml = rowMatches[r].Groups[1].Value;

                var cellMatches = Regex.Matches(
                    rowInnerHtml,
                    @"<(td|th)[^>]*>",
                    RegexOptions.IgnoreCase);

                for (int c = 0; c < cellMatches.Count; c++)
                {
                    string cellTag = cellMatches[c].Value;

                    var classAttrMatch = Regex.Match(
                        cellTag,
                        @"class\s*=\s*[""']([^""']+)[""']",
                        RegexOptions.IgnoreCase);

                    if (classAttrMatch.Success)
                    {
                        string classAttr = classAttrMatch.Groups[1].Value;
                        if (cssClassToColor.TryGetValue(classAttr, out string colorHex))
                        {
                            // Create a style with the background color defined in CSS
                            Style style = worksheet.Cells[startRow + r, c].GetStyle();
                            style.ForegroundColor = ColorTranslator.FromHtml(colorHex);
                            style.Pattern = BackgroundType.Solid;
                            worksheet.Cells[startRow + r, c].SetStyle(style);
                        }
                    }
                }
            }

            // -----------------------------------------------------------------
            // 5. Save the workbook to an Excel file.
            // -----------------------------------------------------------------
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Excel file saved successfully to '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
