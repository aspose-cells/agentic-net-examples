// Title: Convert an HTML file to XLSX while preserving <hr class='page-break'> elements as worksheet page breaks using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads an HTML document, finds <hr> tags marked with a page‑break class, computes the cumulative row index from preceding tables, and adds horizontal page breaks to an Aspose.Cells worksheet before saving as XLSX. | Demonstrate how to load HTML into an Aspose.Cells Workbook with HtmlLoadOptions, then programmatically insert worksheet page breaks at rows derived from the HTML structure in C#.
// Common Searches: aspnet convert html to xlsx preserving page breaks aspocells | c# aspocells add worksheet page break where hr class=page-break appears in html | how to map html table rows to excel rows for pagination using aspocells | preserve html <hr> page break markers when exporting to excel with aspocells .net
// Tags: Aspose.Cells HTML to XLSX with custom page breaks | C# insert horizontal page breaks in worksheet | detect hr page-break class in HTML for Excel export | count table rows for Excel pagination Aspose.Cells | HtmlLoadOptions usage for preserving layout in Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads an HTML file, identifies <hr> elements that act as page‑break markers, counts rows from <table> elements to determine where each break should occur, loads the HTML into an Aspose.Cells workbook, inserts horizontal page breaks at the calculated rows, and saves the result as an XLSX workbook.
class HtmlToExcelWithPageBreaks
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlFilePath = "input.html";

            // Verify that the HTML file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Error: HTML file not found at path '{htmlFilePath}'.");
                return;
            }

            // Load the HTML content into a string for parsing page‑break markers
            string htmlContent = File.ReadAllText(htmlFilePath);

            // List to hold the cumulative row index where each page break should be inserted.
            List<int> pageBreakRowIndices = new List<int>();

            // Keep a running total of rows processed while walking through the document.
            int cumulativeRowCount = 0;

            // Regex to find <table> and <hr> tags in document order.
            Regex tagRegex = new Regex(@"<(table|hr)\b[^>]*>", RegexOptions.IgnoreCase);
            MatchCollection matches = tagRegex.Matches(htmlContent);

            foreach (Match match in matches)
            {
                string tagName = match.Groups[1].Value.ToLowerInvariant();

                if (tagName == "table")
                {
                    // Locate the closing </table> tag to extract the full table markup.
                    int startIdx = match.Index;
                    int endIdx = htmlContent.IndexOf("</table>", startIdx, StringComparison.OrdinalIgnoreCase);
                    if (endIdx >= 0)
                    {
                        string tableHtml = htmlContent.Substring(startIdx, endIdx - startIdx);
                        // Count <tr> elements inside the table (including nested tables).
                        int rowsInTable = Regex.Matches(tableHtml, @"<tr\b", RegexOptions.IgnoreCase).Count;
                        cumulativeRowCount += rowsInTable;
                    }
                }
                else if (tagName == "hr")
                {
                    // Determine whether this <hr> is a page‑break marker.
                    string hrTag = match.Value;
                    if (Regex.IsMatch(hrTag, @"class\s*=\s*[""'][^""']*page-break[^""']*[""']", RegexOptions.IgnoreCase))
                    {
                        // Aspose.Cells adds a horizontal page break *before* the specified row.
                        pageBreakRowIndices.Add(cumulativeRowCount);
                    }
                }
            }

            // Load the HTML into an Aspose.Cells workbook.
            // Use the constructor that accepts a file path and load options.
            Workbook workbook;
            try
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(htmlFilePath, loadOptions);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load HTML into workbook: {loadEx.Message}");
                return;
            }

            // Ensure at least one worksheet was created.
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("No worksheets were created from the HTML content.");
                return;
            }

            // Get the first worksheet (the one created by the HTML import).
            Worksheet sheet = workbook.Worksheets[0];

            // Insert horizontal page breaks at the recorded row indices.
            foreach (int rowIndex in pageBreakRowIndices)
            {
                // Add a page break before the specified row (column index 0 is required by the API).
                sheet.HorizontalPageBreaks.Add(rowIndex, 0);
            }

            // Save the workbook to an Excel file.
            string excelOutputPath = "output.xlsx";
            try
            {
                workbook.Save(excelOutputPath, SaveFormat.Xlsx);
                Console.WriteLine($"HTML converted to Excel with page breaks. Saved to: {excelOutputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save Excel file: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
