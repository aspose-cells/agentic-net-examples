// Title: C# – Convert HTML to Excel with Aspose.Cells and Auto‑Generate a Hyperlinked Table of Contents
// Description: Loads an HTML file into an Aspose.Cells workbook, extracts all <h1>‑<h6> headings, creates a "Table of Contents" worksheet, writes each heading with level‑based indentation, links each entry to the first matching cell in the workbook, and saves the result as an .xlsx file.
// Keywords: Aspose.Cells HTML to Excel conversion | C# generate Excel table of contents | hyperlink headings in Excel | extract h1 h6 tags C# | auto TOC worksheet Aspose.Cells | convert HTML report to Excel | C# regex heading extraction
// Common Searches: how to create a table of contents in Excel from HTML using Aspose.Cells | C# code to add hyperlinks from TOC to HTML headings in Excel | Aspose.Cells load HTML and generate TOC worksheet | extract h1‑h6 tags and map to Excel cells C# | convert HTML report to Excel with clickable navigation
// Developer Intent: Produce an Excel workbook from an HTML source and add a hyperlinked TOC sheet that navigates to each heading.
// Use Cases: Transform marketing or technical reports stored as HTML into Excel workbooks with instant navigation. | Automate documentation publishing where each HTML section appears as a clickable entry in an Excel TOC. | Provide analysts with a consolidated workbook that links summary rows to detailed sections extracted from HTML.
// AI Prompts: Generate C# code using Aspose.Cells to load an HTML file, extract all heading tags, and build a hyperlinked Table of Contents worksheet. | Show how to include page numbers or cell references next to each TOC entry in the sample. | Suggest a strategy for handling duplicate heading texts when creating hyperlinks in the TOC.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells workbook, extracts all <h1>‑<h6> headings, creates a "Table of Contents" worksheet, writes each heading with level‑based indentation, links each entry to the first matching cell in the workbook, and saves the result as an .xlsx file.
class HtmlToExcelWithToc
{
    static void Main()
    {
        try
        {
            // Input HTML file and output Excel file paths
            string htmlPath = "input.html";
            string excelPath = "output.xlsx";

            // Verify that the HTML source file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            // Load the HTML file into a workbook
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Read the raw HTML content for heading extraction
            string htmlContent = File.ReadAllText(htmlPath);

            // Regex to capture <h1> … </h1> through <h6> … </h6>
            Regex headingRegex = new Regex(@"<(h[1-6])[^>]*>(.*?)</\1>", RegexOptions.IgnoreCase);
            MatchCollection headingMatches = headingRegex.Matches(htmlContent);

            // Add a new worksheet that will hold the Table of Contents
            Worksheet tocSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            tocSheet.Name = "Table of Contents";

            int tocRow = 0;

            // Iterate over each heading found in the HTML
            foreach (Match match in headingMatches)
            {
                // Determine heading level (1‑6) and plain text
                string tag = match.Groups[1].Value.ToLower();               // e.g. "h2"
                string headingText = System.Net.WebUtility.HtmlDecode(match.Groups[2].Value.Trim());

                // Indent the entry according to its level for visual hierarchy
                int level = int.Parse(tag.Substring(1)); // 1‑6
                string indentedText = new string(' ', (level - 1) * 4) + headingText;
                tocSheet.Cells[tocRow, 0].PutValue(indentedText);

                // Search for the first cell in the workbook that contains the heading text
                // (skip the TOC sheet itself to avoid self‑reference)
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws == tocSheet) continue;

                    // Find the heading text (case‑insensitive). Use overload with start cell = null.
                    Cell foundCell = ws.Cells.Find(headingText, null, new FindOptions { CaseSensitive = false });
                    if (foundCell != null)
                    {
                        // Create a hyperlink from the TOC entry to the found cell
                        // Hyperlink format: 'SheetName'!A1
                        string hyperlink = $"'{ws.Name}'!{foundCell.Name}";
                        tocSheet.Hyperlinks.Add(tocRow, 0, 1, 1, hyperlink);
                        break;
                    }
                }

                tocRow++;
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(excelPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an Excel file
                workbook.Save(excelPath);
                Console.WriteLine($"Workbook saved successfully to '{excelPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
