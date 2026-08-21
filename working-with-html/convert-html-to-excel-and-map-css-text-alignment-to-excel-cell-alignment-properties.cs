// Title: C# – Convert HTML to Excel and map CSS text‑align to Aspose.Cells cell alignment
// Description: Loads an HTML file into an Aspose.Cells Workbook, reads each cell's HtmlString, extracts the inline CSS text‑align value with a regex, maps it to the appropriate TextAlignmentType enum, applies the horizontal alignment to the cell style, and saves the result as an XLSX workbook. Includes basic error handling for missing files and unsupported alignments.
// Keywords: Aspose.Cells HTML to Excel conversion C# | CSS text-align to Excel alignment | extract inline CSS from HtmlString | TextAlignmentType mapping | preserve cell alignment during HTML‑to‑XLSX | C# regular expression CSS parsing | horizontal alignment in Aspose.Cells | batch HTML to Excel conversion | global web‑to‑spreadsheet export
// Common Searches: Aspose.Cells preserve CSS text-align when converting HTML to Excel | C# extract text-align from cell HtmlString | map CSS alignment values to Aspose.Cells TextAlignmentType | convert HTML tables to XLSX with original alignment | how to keep left center right justification in Excel after HTML import
// Developer Intent: Generate an Excel workbook from an HTML source while retaining the original CSS text‑align styling for each cell using Aspose.Cells in C#.
// Use Cases: Convert web‑based reports or dashboards into Excel files that keep the same left, center, right, justify, or distributed alignment. | Process HTML email templates or scraped web tables and produce XLSX files with matching visual alignment. | Automate bulk conversion of multiple HTML documents to Excel, ensuring inline text‑align styles are reflected in the spreadsheet.
// AI Prompts: Write C# code with Aspose.Cells that loads an HTML file, extracts the text‑align CSS property from each cell's HtmlString, maps it to TextAlignmentType, sets the cell's HorizontalAlignment, and saves the workbook as XLSX. | Show how to extend the CssToAlignmentMap dictionary to support additional CSS values such as 'start', 'end', and 'inherit' and handle them gracefully in Aspose.Cells. | Provide robust error‑handling patterns for missing input files, empty HtmlString values, and unsupported CSS alignment values during HTML‑to‑Excel conversion.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace HtmlToExcelAlignment
{
    // Loads an HTML file into an Aspose.Cells Workbook, reads each cell's HtmlString, extracts the inline CSS text‑align value with a regex, maps it to the appropriate TextAlignmentType enum, applies the horizontal alignment to the cell style, and saves the result as an XLSX workbook. Includes basic error handling for missing files and unsupported alignments.
    class Program
    {
        // Maps CSS text-align values to Aspose.Cells TextAlignmentType enum
        private static readonly Dictionary<string, TextAlignmentType> CssToAlignmentMap = new Dictionary<string, TextAlignmentType>(StringComparer.OrdinalIgnoreCase)
        {
            { "left", TextAlignmentType.Left },
            { "center", TextAlignmentType.Center },
            { "right", TextAlignmentType.Right },
            { "justify", TextAlignmentType.Justify },
            { "distributed", TextAlignmentType.Distributed }
        };

        // Extracts the CSS text-align value from a cell's HTML representation
        private static string? ExtractCssTextAlign(string? html)
        {
            if (string.IsNullOrEmpty(html))
                return null;

            // Look for "text-align: value;" inside a style attribute
            var match = Regex.Match(html, @"text-align\s*:\s*([^;""']+)", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value.Trim() : null;
        }

        static void Main(string[] args)
        {
            try
            {
                // Path to the source HTML file
                string htmlPath = "input.html";

                // Ensure the input file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Input file '{htmlPath}' not found.");
                    return;
                }

                // Load the HTML file into a Workbook
                Workbook workbook = new Workbook(htmlPath, new LoadOptions(LoadFormat.Html));

                // Iterate through all worksheets and cells
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Skip empty cells
                            if (cell == null || cell.Type == CellValueType.IsNull)
                                continue;

                            // Obtain the HTML string of the cell (contains inline CSS)
                            string? cellHtml = cell.HtmlString;

                            // Extract CSS text-align value
                            string? cssAlign = ExtractCssTextAlign(cellHtml);
                            if (string.IsNullOrEmpty(cssAlign))
                                continue; // No alignment defined in CSS

                            // Map CSS alignment to Aspose.Cells alignment enum
                            if (CssToAlignmentMap.TryGetValue(cssAlign, out TextAlignmentType alignment))
                            {
                                // Apply the alignment to the cell's style
                                Style style = cell.GetStyle();
                                style.HorizontalAlignment = alignment;
                                cell.SetStyle(style);
                            }
                        }
                    }
                }

                // Save the workbook as an Excel file
                string excelPath = "output.xlsx";
                workbook.Save(excelPath, SaveFormat.Xlsx);

                Console.WriteLine($"HTML file '{htmlPath}' has been converted to Excel '{excelPath}' with CSS text-align mapped to cell alignment.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
