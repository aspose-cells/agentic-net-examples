// Title: Convert HTML <a> tags to Excel cell hyperlinks while preserving display text using Aspose.Cells for .NET (C#)
// AI Prompts: Load an HTML file into an Aspose.Cells Workbook, locate cells containing <a> elements, replace the cell value with the anchor's inner text, and attach a hyperlink to the same cell via Hyperlinks.Add. | Iterate over the worksheet's used range, apply a regular expression to capture the href URL and link text from each anchor tag, then call sheet.Hyperlinks.Add(row, column, 1, 1, url) to create a single‑cell hyperlink. | Include error handling for missing input files and hyperlink insertion failures, and save the modified workbook as an XLSX file.
// Common Searches: Aspose.Cells C# convert HTML anchor tags to Excel hyperlinks preserving link text | how to add cell hyperlinks after loading HTML with Aspose.Cells | regex extract href and display text from HTML for Aspose.Cells hyperlink creation | load HTML into workbook and replace <a> tags with plain text and hyperlink in .NET
// Tags: HTML to Excel hyperlink conversion Aspose.Cells | Aspose.Cells Hyperlinks.Add single‑cell overload | C# regex extract href display text | preserve link text during HTML import Aspose.Cells | load HTML workbook SaveFormat.Xlsx Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace HtmlAnchorToHyperlinkApp
{
    // The program loads an HTML file into an Aspose.Cells Workbook, scans each used cell for <a> tags, extracts the href URL and inner text using a regular expression, replaces the cell content with the link text, adds a hyperlink to the same cell with Hyperlinks.Add(row, column, 1, 1, url), and saves the result as an XLSX workbook.
    class HtmlAnchorToHyperlink
    {
        static void Main()
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

                // Load the HTML file into a workbook
                Workbook workbook;
                try
                {
                    HtmlLoadOptions loadOptions = new HtmlLoadOptions(LoadFormat.Html);
                    workbook = new Workbook(htmlPath, loadOptions);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load HTML file: {loadEx.Message}");
                    return;
                }

                // Work with the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Regex to capture <a href="...">display text</a>
                Regex anchorRegex = new Regex(
                    @"<a\s+[^>]*href=['""](?<url>[^'""]+)['""][^>]*>(?<text>.*?)</a>",
                    RegexOptions.IgnoreCase);

                // Iterate through all used cells
                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= cells.MaxDataColumn; col++)
                    {
                        string cellText = cells[row, col].StringValue;
                        if (string.IsNullOrEmpty(cellText))
                            continue;

                        Match match = anchorRegex.Match(cellText);
                        if (match.Success)
                        {
                            // Extract URL and display text
                            string url = match.Groups["url"].Value;
                            string displayText = match.Groups["text"].Value;

                            // Replace cell value with the display text
                            cells[row, col].PutValue(displayText);

                            // Add a hyperlink to the same cell (covers a single cell)
                            try
                            {
                                // Use the 5‑argument overload (row, column, rows, columns, url)
                                sheet.Hyperlinks.Add(row, col, 1, 1, url);
                            }
                            catch (Exception hlEx)
                            {
                                Console.WriteLine($"Failed to add hyperlink at {cells[row, col].Name}: {hlEx.Message}");
                            }
                        }
                    }
                }

                // Save the workbook with hyperlinks
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
