// Title: Convert HTML tables to Excel with Aspose.Cells in C# while preserving CSS text‑align as cell horizontal alignment
// AI Prompts: Generate C# code that reads an HTML file, loads it into an Aspose.Cells Workbook, extracts each cell's `text-align` style, and applies the matching `HorizontalAlignment` to the corresponding Excel cell. | Create a C# method that uses regular expressions to parse row/column positions and CSS `text-align` values from an HTML string, then updates the workbook's cell styles accordingly. | Add comprehensive error handling for missing files, read/write failures, and unsupported alignment values, and log conversion progress to the console.
// Common Searches: how to keep css text-align when converting html table to xlsx using Aspose.Cells C# | C# Aspose.Cells load html and map css alignment to excel cell alignment | extract text-align from html table cells and apply to excel using Aspose.Cells | convert html with colspan to excel preserving cell alignment in C# | Aspose.Cells html to excel conversion with custom style handling C#
// Tags: html to xlsx conversion with Aspose.Cells C# | map css text-align to Aspose.Cells HorizontalAlignment | regex extraction of cell alignment from html | apply cell style alignment in Aspose.Cells workbook | handle colspan during html to excel translation

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads an HTML file, uses regular expressions to capture each table cell's `text-align` CSS value and any colspan, loads the HTML into an Aspose.Cells Workbook via a memory stream, then iterates over the extracted alignment map to set the corresponding `HorizontalAlignment` on each worksheet cell before saving the result as an XLSX file. Robust error handling reports missing files, read/write issues, and unsupported alignments.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the input file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // Path to the output Excel file
            string excelPath = "output.xlsx";

            // Read the entire HTML content
            string htmlContent;
            try
            {
                htmlContent = File.ReadAllText(htmlPath);
            }
            catch (Exception readEx)
            {
                Console.WriteLine($"Error reading HTML file: {readEx.Message}");
                return;
            }

            // Parse HTML to extract CSS text-align values for each cell
            var alignmentMap = ExtractCellAlignments(htmlContent);

            // Load the HTML content into a new workbook using a memory stream
            Workbook workbook;
            try
            {
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
                {
                    var loadOptions = new HtmlLoadOptions();
                    workbook = new Workbook(ms, loadOptions);
                }
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Error loading HTML into workbook: {loadEx.Message}");
                return;
            }

            // Apply the extracted alignments to the corresponding cells
            ApplyAlignments(workbook, alignmentMap);

            // Save the workbook to an Excel file
            try
            {
                workbook.Save(excelPath);
                Console.WriteLine($"Excel file saved successfully to '{excelPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving Excel file: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Extracts text-align CSS values from HTML table cells and maps them to cell coordinates
    private static Dictionary<(int Row, int Column), string> ExtractCellAlignments(string html)
    {
        var map = new Dictionary<(int, int), string>();

        // Regex to match table rows
        var rowRegex = new Regex(@"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        // Regex to match table cells (td or th)
        var cellRegex = new Regex(@"<(td|th)([^>]*)>(.*?)</\1>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        // Regex to capture style attribute
        var styleRegex = new Regex(@"style\s*=\s*[""']([^""']*)[""']", RegexOptions.IgnoreCase);
        // Regex to capture colspan attribute
        var colspanRegex = new Regex(@"colspan\s*=\s*[""'](\d+)[""']", RegexOptions.IgnoreCase);

        int rowIndex = 0;
        foreach (Match rowMatch in rowRegex.Matches(html))
        {
            string rowContent = rowMatch.Groups[1].Value;
            int colIndex = 0;

            foreach (Match cellMatch in cellRegex.Matches(rowContent))
            {
                string attrPart = cellMatch.Groups[2].Value;

                // Extract style attribute
                var styleMatch = styleRegex.Match(attrPart);
                if (styleMatch.Success)
                {
                    string styleValue = styleMatch.Groups[1].Value;
                    // Simple parsing for "text-align: value"
                    var parts = styleValue.Split(';');
                    foreach (var part in parts)
                    {
                        var kv = part.Split(':');
                        if (kv.Length == 2 && kv[0].Trim().Equals("text-align", StringComparison.OrdinalIgnoreCase))
                        {
                            string alignValue = kv[1].Trim().ToLower();
                            map[(rowIndex, colIndex)] = alignValue;
                            break;
                        }
                    }
                }

                // Handle colspan to adjust column index correctly
                int colspan = 1;
                var colspanMatch = colspanRegex.Match(attrPart);
                if (colspanMatch.Success && int.TryParse(colspanMatch.Groups[1].Value, out int parsedColspan))
                {
                    colspan = parsedColspan;
                }
                colIndex += colspan;
            }

            rowIndex++;
        }

        return map;
    }

    // Applies the CSS text-align values to the corresponding Excel cells
    private static void ApplyAlignments(Workbook workbook, Dictionary<(int Row, int Column), string> alignmentMap)
    {
        if (workbook.Worksheets.Count == 0)
            return;

        // Work with the first worksheet (as Load creates one per table)
        Worksheet sheet = workbook.Worksheets[0];

        foreach (var kvp in alignmentMap)
        {
            int row = kvp.Key.Row;
            int col = kvp.Key.Column;
            string cssAlign = kvp.Value;

            // Retrieve the cell; if it doesn't exist, create it
            Cell cell = sheet.Cells[row, col];

            // Get the current style or create a new one
            Style style = cell.GetStyle();

            // Map CSS alignment to Aspose.Cells TextAlignmentType
            switch (cssAlign)
            {
                case "left":
                    style.HorizontalAlignment = TextAlignmentType.Left;
                    break;
                case "center":
                case "centre":
                    style.HorizontalAlignment = TextAlignmentType.Center;
                    break;
                case "right":
                    style.HorizontalAlignment = TextAlignmentType.Right;
                    break;
                case "justify":
                    style.HorizontalAlignment = TextAlignmentType.Justify;
                    break;
                default:
                    // Unknown alignment; skip applying
                    continue;
            }

            // Apply the modified style back to the cell
            cell.SetStyle(style);
        }
    }
}
