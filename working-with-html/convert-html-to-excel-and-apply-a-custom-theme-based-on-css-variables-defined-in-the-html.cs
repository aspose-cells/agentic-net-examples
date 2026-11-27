// Title: Convert an HTML file to an Excel workbook and apply a custom theme using CSS variables with Aspose.Cells for .NET
// Common Searches: C# load html file into Aspose.Cells workbook and save as xlsx | How to extract :root CSS variables in C# and use them in Excel theme | Set workbook theme accent colors from CSS variables using Aspose.Cells | Apply custom theme colors to Excel file based on HTML CSS variables in .NET | Aspose.Cells HtmlLoadOptions example for converting HTML to Excel with custom theme
// Tags: aspocells html to xlsx conversion | apply css variables to workbook theme | set themecolortype using c# color dictionary | extract root css variables with regex c# | custom excel theme using aspocells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The sample reads an HTML document, extracts CSS variables defined in a :root block via regular expressions, loads the HTML into an Aspose.Cells Workbook, maps selected variables (primary, secondary, etc.) to the workbook's theme accent slots using SetThemeColor, and saves the result as an XLSX file.
class Program
{
    static void Main(string[] args)
    {
        // Paths – adjust as needed
        string htmlInputPath = "input.html";
        string excelOutputPath = "output.xlsx";

        // Verify HTML input file exists
        if (!File.Exists(htmlInputPath))
        {
            Console.WriteLine($"HTML input file not found: {htmlInputPath}");
            return;
        }

        // Read the HTML content
        string htmlContent = File.ReadAllText(htmlInputPath);

        // Extract CSS variables from the HTML (look for :root { --var: value; ... })
        var cssVarPattern = new Regex(@":root\s*\{([^}]*)\}", RegexOptions.Singleline);
        var varMatch = cssVarPattern.Match(htmlContent);
        var cssVariables = new Dictionary<string, Color>();

        if (varMatch.Success)
        {
            string varsBlock = varMatch.Groups[1].Value;
            var varLinePattern = new Regex(@"--(?<name>[\w-]+)\s*:\s*(?<value>#?[0-9a-fA-F]{3,6})\s*;", RegexOptions.Singleline);
            foreach (Match m in varLinePattern.Matches(varsBlock))
            {
                string name = m.Groups["name"].Value.Trim();
                string value = m.Groups["value"].Value.Trim();

                // Ensure hex color starts with '#'
                if (!value.StartsWith("#"))
                    value = "#" + value;

                try
                {
                    Color color = ColorTranslator.FromHtml(value);
                    cssVariables[name] = color;
                    Console.WriteLine($"Found CSS variable '--{name}': {value}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to parse color value '{value}' for variable '--{name}': {ex.Message}");
                }
            }
        }
        else
        {
            Console.WriteLine("No CSS variables block (:root { ... }) found in the HTML.");
        }

        // Load HTML into a workbook
        Workbook workbook;
        try
        {
            // HtmlLoadOptions does not have ExportActiveWorksheetOnly; default options are sufficient
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            workbook = new Workbook(htmlInputPath, loadOptions);
            Console.WriteLine("HTML successfully loaded into workbook.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading HTML into workbook: {ex.Message}");
            return;
        }

        // Apply a custom theme based on extracted CSS variables
        try
        {
            // Map known CSS variable names to theme color slots
            ApplyThemeColor(workbook, ThemeColorType.Accent1, cssVariables, "primary");
            ApplyThemeColor(workbook, ThemeColorType.Accent2, cssVariables, "secondary");
            ApplyThemeColor(workbook, ThemeColorType.Accent3, cssVariables, "tertiary");
            ApplyThemeColor(workbook, ThemeColorType.Accent4, cssVariables, "quaternary");
            ApplyThemeColor(workbook, ThemeColorType.Accent5, cssVariables, "quinary");
            ApplyThemeColor(workbook, ThemeColorType.Accent6, cssVariables, "senary");

            Console.WriteLine("Custom theme applied based on CSS variables.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error applying custom theme: {ex.Message}");
        }

        // Save the workbook as Excel
        try
        {
            workbook.Save(excelOutputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{excelOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }

    private static void ApplyThemeColor(Workbook workbook, ThemeColorType colorType,
        Dictionary<string, Color> variables, string cssVarName)
    {
        if (variables.TryGetValue(cssVarName, out Color color))
        {
            workbook.SetThemeColor(colorType, color);
            Console.WriteLine($"Applied CSS variable '--{cssVarName}' to theme slot {colorType}.");
        }
        else
        {
            Console.WriteLine($"CSS variable '--{cssVarName}' not found; leaving theme slot {colorType} unchanged.");
        }
    }
}
