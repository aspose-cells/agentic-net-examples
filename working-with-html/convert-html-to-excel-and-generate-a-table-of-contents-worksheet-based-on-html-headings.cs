// Title: Convert HTML file to Excel workbook and create an indented Table of Contents sheet from H1‑H6 headings using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads an HTML file, imports it into an Aspose.Cells Workbook, extracts every <h1>‑<h6> tag, and populates a new "Table of Contents" worksheet with rows indented according to heading level. | Enhance the previous solution to apply bold formatting to level‑1 headings and automatically adjust column widths for the TOC sheet. | Create a reusable method `BuildWorkbookWithToc(string htmlContent)` that returns a Workbook containing the original HTML data and a formatted TOC worksheet, then demonstrate saving it as an .xlsx file.
// Common Searches: asp.net convert html page to xlsx and generate toc based on heading hierarchy | c# Aspose.Cells load html and create table of contents worksheet from h2 h3 tags | how to indent rows in Excel according to html heading level using Aspose.Cells | extract headings from html and write them to separate columns in Excel with Aspose.Cells
// Tags: import html via Aspose.Cells LoadOptions | capture heading tags with C# regular expressions | build Excel TOC worksheet from extracted headings | apply indentation to rows based on heading depth | style top‑level headings bold in Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program reads an HTML file, loads it into an Aspose.Cells Workbook using LoadOptions, extracts all <h1>‑<h6> headings via a regular expression, creates a new worksheet named "Table of Contents", writes each heading into a row with indentation that reflects its level, applies bold formatting to level‑1 headings, and saves the result as an XLSX file.
class HtmlToExcelWithToc
{
    static void Main()
    {
        try
        {
            // Path to the input HTML file
            string htmlPath = "input.html";

            // Verify that the HTML file exists to avoid FileNotFoundException
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // Load the HTML content from the file
            string htmlContent = File.ReadAllText(htmlPath);

            // Load HTML into a workbook using LoadOptions for HTML format
            Workbook workbook;
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                workbook = new Workbook(ms, loadOptions);
            }

            // Add a new worksheet that will serve as the Table of Contents
            Worksheet tocSheet = workbook.Worksheets.Add("Table of Contents");

            // Extract heading elements (h1‑h6) from the HTML using regular expressions
            var headingPattern = new Regex(@"<(h[1-6])[^>]*>(.*?)</\1>", RegexOptions.IgnoreCase);
            var matches = headingPattern.Matches(htmlContent);

            if (matches.Count > 0)
            {
                int rowIndex = 0; // start at the first row of the TOC sheet

                foreach (Match match in matches)
                {
                    // Determine heading level (1‑6) from the tag name (e.g., h2 -> 2)
                    string tagName = match.Groups[1].Value.ToLower(); // e.g., "h2"
                    int level = int.Parse(tagName.Substring(1));

                    // Extract and clean the heading text
                    string headingText = Regex.Replace(match.Groups[2].Value, @"\s+", " ").Trim();

                    // Create a cell for the heading text, indent based on level
                    Cell cell = tocSheet.Cells[rowIndex, level - 1];
                    cell.PutValue(headingText);

                    // Optional: style the heading (bold for top‑level headings)
                    Style style = cell.GetStyle();
                    if (level == 1)
                    {
                        style.Font.IsBold = true;
                    }
                    cell.SetStyle(style);

                    rowIndex++;
                }
            }

            // Save the workbook to an Excel file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
