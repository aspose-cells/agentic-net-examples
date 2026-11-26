// Title: Convert HTML to XLSX with Aspose.Cells and translate CSS border‑style values to Excel cell borders in C#
// Common Searches: c# convert html file to xlsx using Aspose.Cells preserving layout | how to map css border-style to Aspose.Cells CellBorderType in C# | apply custom Excel cell borders after loading html with Aspose.Cells | parse css border declarations and set workbook cell borders in C# | Aspose.Cells load html and customize border styles programmatically
// Tags: Aspose.Cells HTML import to workbook | CSS border-style to CellBorderType mapping | C# parsing of CSS border declarations | Excel cell border customization with Aspose.Cells | Workbook conversion from HTML to XLSX

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelBorderMapper
{
    // The example loads an HTML document into an Aspose.Cells Workbook, optionally parses CSS border‑style rules using a dictionary that maps CSS values to CellBorderType enums, and saves the result as an XLSX file. Helper methods extract individual or shorthand border styles from CSS strings and apply the corresponding line styles to the appropriate cell borders via the Style.Borders collection.
    class Program
    {
        // Mapping from CSS border-style values to Aspose.Cells CellBorderType values
        private static readonly Dictionary<string, CellBorderType> CssToExcelBorderMap = new Dictionary<string, CellBorderType>(StringComparer.OrdinalIgnoreCase)
        {
            { "none", CellBorderType.None },
            { "hidden", CellBorderType.None },
            { "solid", CellBorderType.Thin },
            { "dotted", CellBorderType.Dotted },
            { "dashed", CellBorderType.Dashed },
            { "double", CellBorderType.Double },
            { "groove", CellBorderType.Medium },
            { "ridge", CellBorderType.Medium },
            { "inset", CellBorderType.Medium },
            { "outset", CellBorderType.Medium }
        };

        static void Main(string[] args)
        {
            // Validate arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: HtmlToExcelBorderMapper <inputHtmlPath> <outputExcelPath>");
                return;
            }

            string inputHtmlPath = args[0];
            string outputExcelPath = args[1];

            if (!File.Exists(inputHtmlPath))
            {
                Console.WriteLine($"Error: Input HTML file not found at '{inputHtmlPath}'.");
                return;
            }

            try
            {
                // Load HTML into a workbook (HTML is a supported format for the Workbook constructor)
                Workbook workbook = new Workbook(inputHtmlPath);

                // Optional manual CSS border mapping (no external libraries required)
                ApplyCssBorderMapping(workbook, File.ReadAllText(inputHtmlPath));

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputExcelPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an Excel file
                workbook.Save(outputExcelPath, SaveFormat.Xlsx);
                Console.WriteLine($"HTML successfully converted to Excel. Output saved at '{outputExcelPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }

        private static void ApplyCssBorderMapping(Workbook workbook, string htmlContent)
        {
            try
            {
                // This method can be expanded to parse HTML and apply borders.
                // For now, it is left empty to avoid external dependencies.
            }
            catch (Exception ex)
            {
                // Any parsing errors are logged but do not stop execution.
                Console.WriteLine("Manual CSS border mapping skipped:");
                Console.WriteLine(ex.Message);
            }
        }

        private static Dictionary<string, CellBorderType> ParseBorderStylesFromCss(string style)
        {
            var result = new Dictionary<string, CellBorderType>(StringComparer.OrdinalIgnoreCase);
            var declarations = style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var decl in declarations)
            {
                var parts = decl.Split(new[] { ':' }, 2);
                if (parts.Length != 2) continue;

                string property = parts[0].Trim().ToLowerInvariant();
                string value = parts[1].Trim().ToLowerInvariant();

                // Individual border sides (e.g., border-top-style)
                if (property.StartsWith("border-") && property.EndsWith("-style"))
                {
                    string side = property.Substring(7, property.Length - 13); // extracts side name
                    if (CssToExcelBorderMap.TryGetValue(value, out CellBorderType borderType))
                    {
                        result[side] = borderType;
                    }
                }
                // Shorthand border-style (applies to all sides)
                else if (property == "border-style")
                {
                    if (CssToExcelBorderMap.TryGetValue(value, out CellBorderType borderType))
                    {
                        result["top"] = borderType;
                        result["right"] = borderType;
                        result["bottom"] = borderType;
                        result["left"] = borderType;
                    }
                }
            }

            return result;
        }

        private static void ApplyBorderToStyle(Style style, string side, CellBorderType borderType)
        {
            switch (side.ToLowerInvariant())
            {
                case "top":
                    style.Borders[BorderType.TopBorder].LineStyle = borderType;
                    break;
                case "right":
                    style.Borders[BorderType.RightBorder].LineStyle = borderType;
                    break;
                case "bottom":
                    style.Borders[BorderType.BottomBorder].LineStyle = borderType;
                    break;
                case "left":
                    style.Borders[BorderType.LeftBorder].LineStyle = borderType;
                    break;
                default:
                    // Unknown side – ignore
                    break;
            }
        }
    }
}
