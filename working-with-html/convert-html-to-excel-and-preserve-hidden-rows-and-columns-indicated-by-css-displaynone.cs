// Title: C# – Convert HTML to Excel with Aspose.Cells while Keeping CSS‑Hidden Rows & Columns
// Description: Loads an HTML file into an Aspose.Cells workbook, parses the markup to find rows or cells styled with display:none, applies HideRow/HideColumn, and saves the result as an XLSX so hidden elements stay hidden.
// Keywords: Aspose.Cells | HTML to Excel | C# convert HTML | preserve hidden rows | preserve hidden columns | display:none | HideRow | HideColumn | Workbook | HTML parsing
// Common Searches: Aspose.Cells keep hidden rows when converting HTML to Excel | C# detect CSS display:none and hide Excel columns | HTML table to XLSX preserving hidden columns
// Developer Intent: Convert an HTML document to an Excel workbook and retain any rows or columns that are hidden via CSS display:none.
// Use Cases: Export web‑based dashboards to Excel while maintaining role‑based hidden rows. | Migrate legacy HTML reports to XLSX files without exposing calculation columns. | Generate Excel templates from HTML where hidden cells store intermediate data.
// AI Prompts: Show a C# Aspose.Cells example that reads an HTML file, detects display:none on rows and cells, and hides the corresponding rows/columns in the Excel output. | Explain how to extend the regex logic to read display:none rules from an external CSS file during HTML‑to‑Excel conversion. | Suggest performance‑friendly ways to handle colspan/rowspan while preserving hidden rows and columns in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells workbook, parses the markup to find rows or cells styled with display:none, applies HideRow/HideColumn, and saves the result as an XLSX so hidden elements stay hidden.
class HtmlToExcelWithHiddenRowsColumns
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Verify that the HTML file exists
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: HTML file '{htmlPath}' not found.");
            return;
        }

        try
        {
            // Load the HTML file into an Aspose.Cells workbook
            Workbook workbook = new Workbook(htmlPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Read the HTML content for custom parsing (to detect hidden rows/columns)
            string htmlContent = File.ReadAllText(htmlPath);

            // Keep track of hidden row indices (0‑based)
            var hiddenRows = new HashSet<int>();

            // Keep track of hidden column indices (0‑based)
            var hiddenColumns = new HashSet<int>();

            // Regex to match <tr ...>...</tr>
            var trRegex = new Regex(@"<tr\b([^>]*)>(.*?)</tr>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var trMatches = trRegex.Matches(htmlContent);

            for (int rowIndex = 0; rowIndex < trMatches.Count; rowIndex++)
            {
                var trMatch = trMatches[rowIndex];
                string trAttributes = trMatch.Groups[1].Value;

                // Check if the entire row is hidden via style attribute
                var styleMatch = Regex.Match(trAttributes, @"style\s*=\s*[""']([^""']*)[""']", RegexOptions.IgnoreCase);
                if (styleMatch.Success && styleMatch.Groups[1].Value.IndexOf("display:none", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    hiddenRows.Add(rowIndex);
                    continue; // Skip cell processing for a fully hidden row
                }

                // Regex to match <th ...> or <td ...> within the current <tr>
                var cellRegex = new Regex(@"<(th|td)\b([^>]*)>", RegexOptions.IgnoreCase);
                var cellMatches = cellRegex.Matches(trMatch.Groups[2].Value);

                int colPos = 0;
                foreach (Match cellMatch in cellMatches)
                {
                    string cellAttributes = cellMatch.Groups[2].Value;

                    // Determine colspan (default = 1)
                    int colspan = 1;
                    var colspanMatch = Regex.Match(cellAttributes, @"colspan\s*=\s*[""'](\d+)[""']", RegexOptions.IgnoreCase);
                    if (colspanMatch.Success && int.TryParse(colspanMatch.Groups[1].Value, out int cs))
                        colspan = cs;

                    // Check cell style for display:none
                    var cellStyleMatch = Regex.Match(cellAttributes, @"style\s*=\s*[""']([^""']*)[""']", RegexOptions.IgnoreCase);
                    if (cellStyleMatch.Success && cellStyleMatch.Groups[1].Value.IndexOf("display:none", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Mark all columns spanned by this cell as hidden
                        for (int i = 0; i < colspan; i++)
                            hiddenColumns.Add(colPos + i);
                    }

                    colPos += colspan;
                }
            }

            // Apply hidden rows to the worksheet
            foreach (int rowIdx in hiddenRows)
                sheet.Cells.HideRow(rowIdx);

            // Apply hidden columns to the worksheet
            foreach (int colIdx in hiddenColumns)
                sheet.Cells.HideColumn(colIdx);

            // Save the workbook as an Excel file
            string excelPath = "output.xlsx";
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML converted to Excel. Hidden rows/columns preserved in '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
