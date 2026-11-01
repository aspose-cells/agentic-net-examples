// Title: Convert HTML to XLSX and apply a custom workbook theme using CSS variables with Aspose.Cells for .NET
// AI Prompts: Write C# code that reads an HTML file, parses the :root CSS variables with regular expressions, and loads the HTML into an Aspose.Cells Workbook. | Create a helper method that converts hex color strings from the extracted CSS variables to System.Drawing.Color and assigns those colors to the workbook's default font and background. | Use a Style and StyleFlag to apply the colors to every used cell in each worksheet, then save the workbook as an XLSX file.
// Common Searches: how to load HTML into Aspose.Cells workbook in C# | extract :root CSS variables from an HTML document using C# regex | set workbook theme colors from hex values with Aspose.Cells | apply a style to the entire used range of a worksheet in Aspose.Cells
// Tags: Aspose.Cells HTML to XLSX conversion | regex parse root CSS variables C# | hex color to System.Drawing.Color conversion | apply default workbook style Aspose.Cells | StyleFlag apply to used range Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads an HTML file, extracts CSS variables defined in a :root block via regex, loads the HTML into an Aspose.Cells Workbook, converts the hex colors to System.Drawing.Color, applies those colors to the workbook's default and cell styles across all used cells, and saves the result as an XLSX workbook.
class HtmlToExcelWithTheme
{
    static void Main()
    {
        try
        {
            // ---------- Load HTML content ----------
            string htmlPath = "input.html";
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: Input file '{htmlPath}' not found.");
                return;
            }

            string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);

            // ---------- Extract CSS variables ----------
            var cssVariables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Find all <style> blocks
            foreach (Match styleMatch in Regex.Matches(htmlContent, @"<style[^>]*>(.*?)</style>", RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                string styleText = styleMatch.Groups[1].Value;

                // Capture the :root selector block
                var rootMatch = Regex.Match(styleText, @":root\s*\{([^}]*)\}", RegexOptions.Singleline);
                if (rootMatch.Success)
                {
                    string varsBlock = rootMatch.Groups[1].Value;

                    // Match each variable definition: --name: value;
                    foreach (Match varMatch in Regex.Matches(varsBlock, @"(--[\w-]+)\s*:\s*([^;]+);"))
                    {
                        string varName = varMatch.Groups[1].Value.Trim();
                        string varValue = varMatch.Groups[2].Value.Trim();
                        cssVariables[varName] = varValue;
                    }
                }
            }

            // ---------- Load HTML into Aspose.Cells Workbook ----------
            var loadOptions = new HtmlLoadOptions(); // default options

            using (var htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                Workbook workbook;
                try
                {
                    workbook = new Workbook(htmlStream, loadOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading HTML into workbook: {ex.Message}");
                    return;
                }

                // ---------- Helper: Convert hex color to System.Drawing.Color ----------
                Color HexToColor(string hex)
                {
                    if (string.IsNullOrWhiteSpace(hex))
                        return Color.Black;

                    hex = hex.TrimStart('#');

                    // Expand short notation (#abc) to full (#aabbcc)
                    if (hex.Length == 3)
                    {
                        hex = string.Concat(hex[0], hex[0], hex[1], hex[1], hex[2], hex[2]);
                    }

                    // Parse RGB or ARGB
                    if (hex.Length == 6)
                    {
                        int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                        int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                        int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                        return Color.FromArgb(r, g, b);
                    }
                    if (hex.Length == 8)
                    {
                        int a = Convert.ToInt32(hex.Substring(0, 2), 16);
                        int r = Convert.ToInt32(hex.Substring(2, 2), 16);
                        int g = Convert.ToInt32(hex.Substring(4, 2), 16);
                        int b = Convert.ToInt32(hex.Substring(6, 2), 16);
                        return Color.FromArgb(a, r, g, b);
                    }

                    // Invalid format – fallback to black
                    return Color.Black;
                }

                // Determine theme colors from CSS variables (fallbacks provided)
                Color primaryColor = cssVariables.TryGetValue("--primary-color", out var primaryHex)
                    ? HexToColor(primaryHex)
                    : Color.Black;

                Color secondaryColor = cssVariables.TryGetValue("--secondary-color", out var secondaryHex)
                    ? HexToColor(secondaryHex)
                    : Color.White;

                // Apply colors to the workbook's default style
                Style defaultStyle = workbook.DefaultStyle;
                defaultStyle.Font.Color = primaryColor;
                defaultStyle.ForegroundColor = secondaryColor;
                defaultStyle.Pattern = BackgroundType.Solid;

                // Create a reusable style for used cells
                Style cellStyle = workbook.CreateStyle();
                cellStyle.Font.Color = primaryColor;
                cellStyle.ForegroundColor = secondaryColor;
                cellStyle.Pattern = BackgroundType.Solid;

                // Prepare a StyleFlag that applies all style attributes
                StyleFlag flag = new StyleFlag { All = true };

                // Apply the style to all used cells for a uniform appearance
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    var usedRange = sheet.Cells.MaxDisplayRange;
                    if (usedRange != null)
                    {
                        usedRange.ApplyStyle(cellStyle, flag);
                    }
                }

                // ---------- Save the workbook ----------
                string outputPath = "output.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"HTML converted to Excel with custom theme applied successfully. Output: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
