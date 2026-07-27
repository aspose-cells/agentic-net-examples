// Title: Export Aspose.Cells Workbook to HTML and Extract Hyperlink URLs in C#
// Description: Creates a workbook, adds cells with hyperlinks, saves it as HTML with links opening in a new tab, reads the HTML file, extracts all href values using a regular expression, and prints the URLs to the console.
// Keywords: Aspose.Cells HTML export | C# hyperlink extraction | regex href parsing | HtmlSaveOptions LinkTargetType Blank | .NET spreadsheet to web page | parse generated HTML for URLs | Aspose.Cells link audit | C# HTML parsing | SEO link extraction from Excel
// Common Searches: how to save Aspose.Cells workbook as HTML in C# | extract href attributes from generated HTML using regex | retrieve all hyperlinks from Aspose.Cells HTML output | C# parse HTML file for URLs | Aspose.Cells export with links opening in new tab
// Developer Intent: Generate an HTML version of a workbook and programmatically collect every hyperlink URL it contains.
// Use Cases: Publish a spreadsheet as a web page and verify that all embedded links are correct. | Automate SEO audits by gathering URLs from a report generated with Aspose.Cells. | Create a link‑checking script that validates destinations in exported HTML files. | Integrate hyperlink extraction into a CI pipeline to ensure documentation links remain functional.
// AI Prompts: Replace the regex logic with HtmlAgilityPack to safely parse href attributes from the saved HTML. | Write code that saves the extracted URLs to a CSV or JSON file instead of printing them. | Show how to configure HtmlSaveOptions to embed CSS inline while keeping hyperlink functionality. | Demonstrate how to filter extracted URLs by domain (e.g., only external links).

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlLinkExtraction
{
    // Creates a workbook, adds cells with hyperlinks, saves it as HTML with links opening in a new tab, reads the HTML file, extracts all href values using a regular expression, and prints the URLs to the console.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Add sample data and hyperlinks to the worksheet
                sheet.Cells["A1"].PutValue("Google");
                sheet.Hyperlinks.Add("A1", 1, 1, "https://www.google.com");

                sheet.Cells["A2"].PutValue("Aspose");
                sheet.Hyperlinks.Add("A2", 1, 1, "https://www.aspose.com");

                sheet.Cells["A3"].PutValue("GitHub");
                sheet.Hyperlinks.Add("A3", 1, 1, "https://github.com");

                // 3. Configure HTML save options (open links in new tab)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    LinkTargetType = HtmlLinkTargetType.Blank
                };

                // 4. Save the workbook as an HTML file
                string htmlPath = "WorkbookOutput.html";
                workbook.Save(htmlPath, htmlOptions);

                // 5. Ensure the HTML file was created
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"HTML file not found: {htmlPath}");
                    return;
                }

                // 6. Read the HTML content and extract href attributes using a regex
                string htmlContent = File.ReadAllText(htmlPath);
                List<string> extractedUrls = new List<string>();
                foreach (Match match in Regex.Matches(htmlContent, @"href\s*=\s*[""']([^""']+)[""']",
                    RegexOptions.IgnoreCase))
                {
                    string url = match.Groups[1].Value;
                    if (!string.IsNullOrEmpty(url))
                    {
                        extractedUrls.Add(url);
                    }
                }

                // 7. Output the extracted URLs
                Console.WriteLine("Extracted hyperlink URLs from HTML:");
                foreach (string url in extractedUrls)
                {
                    Console.WriteLine(url);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
