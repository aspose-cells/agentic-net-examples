// Title: Convert HTML to XLSX with Aspose.Cells and apply a custom theme using CSS variables in C#
// AI Prompts: Write C# code that loads an HTML document into an Aspose.Cells Workbook, reads CSS custom properties, and uses them to style the workbook's default and header rows. | Show how to parse CSS variable definitions from a <style> block and translate the values into System.Drawing.Color objects for Aspose.Cells styling. | Demonstrate adding column width adjustments based on numeric CSS variables before saving the workbook as an XLSX file.
// Common Searches: aspocells html to xlsx conversion preserving css styling | c# extract css custom properties from html for excel theming | apply html css colors to aspocells workbook default style | set header row formatting in excel using css variables from html | convert html file to excel and use css variables for font and background in c#
// Tags: html to xlsx conversion aspocells | css variable extraction c# | workbook default style customization aspocells | header row theming aspocells | css color to system.drawing.color mapping

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using Aspose.Cells;

// The example reads an HTML file, loads it into an Aspose.Cells Workbook, extracts CSS custom property definitions via regex, converts the CSS colors to System.Drawing.Color, applies those values to the workbook's default style and a styled header row, and saves the result as an XLSX file.
class HtmlToExcelWithCustomTheme
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Ensure the HTML file exists to avoid FileNotFoundException
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // Read the entire HTML content
            string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);

            // --------------------------------------------------------------------
            // 1. Load HTML into an Aspose.Cells Workbook
            // --------------------------------------------------------------------
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // ----------------------------------------------------------------
                // 2. Extract CSS variable definitions from the HTML
                // ----------------------------------------------------------------
                Dictionary<string, string> cssVariables = ExtractCssVariables(htmlContent);

                // ----------------------------------------------------------------
                // 3. Apply a simple custom "theme" using the extracted variables
                // ----------------------------------------------------------------
                ApplyCustomTheme(workbook, cssVariables);

                // ----------------------------------------------------------------
                // 4. Save the resulting Excel file
                // ----------------------------------------------------------------
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Excel file saved to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    private static Dictionary<string, string> ExtractCssVariables(string html)
    {
        var variables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Regex to capture CSS variable definitions inside any <style> block.
        // It looks for patterns like "--var-name: value;" possibly surrounded by whitespace.
        string pattern = @"--(?<name>[\w-]+)\s*:\s*(?<value>[^;]+);";
        foreach (Match match in Regex.Matches(html, pattern))
        {
            string name = match.Groups["name"].Value.Trim();
            string value = match.Groups["value"].Value.Trim();
            if (!variables.ContainsKey(name))
            {
                variables.Add(name, value);
            }
        }

        return variables;
    }

    private static void ApplyCustomTheme(Workbook workbook, Dictionary<string, string> cssVars)
    {
        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Use the workbook's default style as the base for whole‑sheet styling
        Style defaultStyle = workbook.DefaultStyle;

        // Helper to convert CSS color strings to System.Drawing.Color
        Color? ParseColor(string cssColor)
        {
            try
            {
                return ColorTranslator.FromHtml(cssColor);
            }
            catch
            {
                return null;
            }
        }

        // Apply background color to the whole sheet if defined
        if (cssVars.TryGetValue("primary-bg", out string bgColorStr))
        {
            var bgColor = ParseColor(bgColorStr);
            if (bgColor.HasValue)
            {
                defaultStyle.ForegroundColor = bgColor.Value;
                defaultStyle.Pattern = BackgroundType.Solid;
            }
        }

        // Apply font color to the whole sheet if defined
        if (cssVars.TryGetValue("primary-fg", out string fgColorStr))
        {
            var fgColor = ParseColor(fgColorStr);
            if (fgColor.HasValue)
            {
                defaultStyle.Font.Color = fgColor.Value;
            }
        }

        // Apply font family if defined
        if (cssVars.TryGetValue("font-family", out string fontFamily))
        {
            defaultStyle.Font.Name = fontFamily;
        }

        // Apply font size if defined (convert to int as Aspose.Cells expects an integer)
        if (cssVars.TryGetValue("font-size", out string fontSizeStr) &&
            double.TryParse(fontSizeStr.Replace("px", ""), out double fontSize))
        {
            defaultStyle.Font.Size = (int)fontSize;
        }

        // The workbook's default style is automatically used; no need to reassign.

        // Apply header-specific styling (first row) if variables exist
        if (cssVars.Count > 0)
        {
            // Determine the used range to know how many columns the header spans
            var usedRange = sheet.Cells.MaxDisplayRange;
            int lastColumn = usedRange.ColumnCount - 1;

            // Create a style for the header row
            Style headerStyle = workbook.CreateStyle();

            if (cssVars.TryGetValue("header-bg", out string headerBgStr))
            {
                var headerBg = ParseColor(headerBgStr);
                if (headerBg.HasValue)
                {
                    headerStyle.ForegroundColor = headerBg.Value;
                    headerStyle.Pattern = BackgroundType.Solid;
                }
            }

            if (cssVars.TryGetValue("header-fg", out string headerFgStr))
            {
                var headerFg = ParseColor(headerFgStr);
                if (headerFg.HasValue)
                {
                    headerStyle.Font.Color = headerFg.Value;
                }
            }

            // Optional: make header bold
            headerStyle.Font.IsBold = true;

            // Apply the header style to the first row across all used columns
            for (int col = 0; col <= lastColumn; col++)
            {
                sheet.Cells[0, col].SetStyle(headerStyle);
            }
        }
    }
}
