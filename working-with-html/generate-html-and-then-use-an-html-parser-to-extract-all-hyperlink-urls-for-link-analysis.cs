// Title: C# – Save Aspose.Cells workbook as HTML and extract every hyperlink URL
// Description: The sample creates a new Aspose.Cells workbook, inserts cells with hyperlinks, and saves the sheet as an HTML file using HtmlSaveOptions (blank target and full‑path links). It then reads the generated HTML, parses all <a href> attributes with a regular expression, collects the URLs into a list, and prints them for link‑analysis or validation purposes.
// Keywords: Aspose.Cells C# HTML export | hyperlink extraction C# | regex href parsing | HtmlSaveOptions LinkTargetType.Blank | full path links Aspose HTML | HTML link analysis | C# workbook to HTML | extract URLs from HTML | Aspose.Cells hyperlink example | GitHub Aspose.Cells code
// Common Searches: save Aspose.Cells workbook as HTML C# | extract href links from Aspose generated HTML | C# regex to get all URLs from HTML file | how to parse hyperlinks in Aspose.Cells HTML output | link extraction for SEO from spreadsheet HTML
// Developer Intent: Generate an HTML representation of a spreadsheet and retrieve every hyperlink URL contained in that HTML for further processing.
// Use Cases: Verify that spreadsheet hyperlinks are correctly exported before publishing a web report. | Gather external URLs from multiple workbooks for bulk SEO or compliance audits. | Feed extracted links into a crawler or validation tool to detect broken or malicious URLs. | Create a link‑inventory for marketing analytics from automatically generated HTML reports.
// AI Prompts: Show how to replace the regex with HtmlAgilityPack for more reliable href extraction. | Provide code to export the workbook with relative links while still being able to parse them. | Explain how to deduplicate URLs and safely handle malformed href attributes during parsing. | Generate a PowerShell script that runs this C# program across a folder of workbooks and aggregates all extracted links.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlLinkExtraction
{
    // The sample creates a new Aspose.Cells workbook, inserts cells with hyperlinks, and saves the sheet as an HTML file using HtmlSaveOptions (blank target and full‑path links). It then reads the generated HTML, parses all <a href> attributes with a regular expression, collects the URLs into a list, and prints them for link‑analysis or validation purposes.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a workbook and add sample hyperlinks ----------
                Workbook workbook = new Workbook();                         // create a new workbook
                Worksheet sheet = workbook.Worksheets[0];                  // get the first worksheet

                // Add sample text and hyperlinks
                sheet.Cells["A1"].PutValue("Google");
                sheet.Hyperlinks.Add("A1", 1, 1, "https://www.google.com");

                sheet.Cells["A2"].PutValue("Aspose");
                sheet.Hyperlinks.Add("A2", 1, 1, "https://www.aspose.com");

                sheet.Cells["A3"].PutValue("GitHub");
                sheet.Hyperlinks.Add("A3", 1, 1, "https://github.com");

                // ---------- Save the workbook as HTML ----------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    LinkTargetType = HtmlLinkTargetType.Blank, // open links in a new tab/window
                    IsFullPathLink = true                       // use full path links in the HTML
                };
                string htmlPath = "output.html";
                workbook.Save(htmlPath, htmlOptions); // save workbook to HTML

                // ---------- Verify the HTML file exists ----------
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: HTML file '{htmlPath}' was not created.");
                    return;
                }

                // ---------- Parse the generated HTML and extract hyperlink URLs ----------
                string htmlContent = File.ReadAllText(htmlPath);
                // Simple regex to capture href values within <a> tags
                Regex hrefRegex = new Regex(@"<a[^>]+href\s*=\s*[""'](?<url>[^""'>]+)[""']", RegexOptions.IgnoreCase);
                MatchCollection matches = hrefRegex.Matches(htmlContent);

                List<string> extractedUrls = new List<string>();
                foreach (Match match in matches)
                {
                    string url = match.Groups["url"].Value;
                    if (!string.IsNullOrEmpty(url))
                    {
                        extractedUrls.Add(url);
                    }
                }

                // ---------- Output the extracted URLs ----------
                Console.WriteLine("Extracted hyperlink URLs from HTML:");
                foreach (string url in extractedUrls)
                {
                    Console.WriteLine(url);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
