using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlImport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source HTML file
                string htmlPath = "input.html";

                // Verify that the HTML file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: File \"{htmlPath}\" not found.");
                    return;
                }

                // Load the HTML file into a workbook
                HtmlLoadOptions loadOptions = new HtmlLoadOptions
                {
                    // Treat any formula‑like text as plain text
                    HasFormula = false
                };
                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // Work with the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Regex to capture <a href="url">display text</a>
                Regex anchorRegex = new Regex(
                    @"<a\s+[^>]*href\s*=\s*['""](?<url>[^'""]+)['""][^>]*>(?<text>.*?)</a>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                // Iterate through all used cells
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        string cellValue = cells[row, col].StringValue;
                        if (string.IsNullOrEmpty(cellValue))
                            continue;

                        Match match = anchorRegex.Match(cellValue);
                        if (match.Success)
                        {
                            // Extract URL and display text
                            string url = match.Groups["url"].Value;
                            string displayText = match.Groups["text"].Value;

                            // Replace cell content with the display text
                            cells[row, col].PutValue(displayText);

                            // Add a hyperlink to the same cell (single‑cell range)
                            // Using overload with 5 parameters (no screen tip) for compatibility
                            sheet.Hyperlinks.Add(row, col, 1, 1, url);
                        }
                    }
                }

                // Save the workbook to Excel format
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}