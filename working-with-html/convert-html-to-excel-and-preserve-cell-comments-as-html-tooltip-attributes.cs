// Title: Convert an HTML table to an XLSX file while preserving HTML title attributes as cell comments using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML file, extracts table rows and cells, writes the cell values to an Aspose.Cells workbook, and creates a comment for each cell from its title attribute. | Show how to use Aspose.Cells to add cell comments based on HTML tooltip attributes when converting an HTML table to an Excel worksheet. | Provide a C# example that parses <td> and <th> elements with regular expressions, saves the data to a .xlsx file, and attaches the title attribute as a comment.
// Common Searches: aspocells c# convert html table to xlsx with cell comments from title attribute | how to preserve html tooltip as Excel comment using Aspose.Cells | c# read html file and export to excel while keeping tooltip notes | extract title attribute from td and add as comment in Aspose.Cells workbook | convert html table to excel preserving cell notes .net
// Tags: html-to-xlsx conversion with Aspose.Cells | extract title attribute to Excel comment | c# regex parsing html for Aspose.Cells export | add cell comments from HTML tooltip Aspose.Cells | save workbook as xlsx with cell notes

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example reads an HTML file, uses regular expressions to locate table rows and cells, writes each cell's plain text into a new Aspose.Cells workbook, extracts any title attribute from the HTML cell and adds it as a comment on the corresponding Excel cell, then saves the workbook as an .xlsx file.
class HtmlToExcelWithComments
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
                Console.WriteLine($"Input file '{htmlPath}' not found.");
                return;
            }

            // Load the HTML content into a string
            string htmlContent = File.ReadAllText(htmlPath);

            // Create a new workbook (Aspose.Cells)
            var workbook = new Workbook();

            // Use the first worksheet
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Regular expression to find table rows
            var rowMatches = Regex.Matches(htmlContent, @"<tr[^>]*>(.*?)</tr>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            if (rowMatches.Count == 0)
            {
                Console.WriteLine("No table rows found in the HTML document.");
                return;
            }

            int rowIndex = 0;
            foreach (Match rowMatch in rowMatches)
            {
                // Regular expression to find cells (td or th) within the current row
                var cellMatches = Regex.Matches(rowMatch.Groups[1].Value, @"<(td|th)([^>]*)>(.*?)</\1>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                int colIndex = 0;
                foreach (Match cellMatch in cellMatches)
                {
                    // Extract inner text (strip any nested tags)
                    string innerHtml = cellMatch.Groups[3].Value;
                    string cellText = Regex.Replace(innerHtml, "<.*?>", string.Empty).Trim();

                    // Write the text to the corresponding Excel cell
                    var cell = sheet.Cells[rowIndex, colIndex];
                    cell.PutValue(cellText);

                    // Extract the title attribute if present and add as a comment
                    string attributes = cellMatch.Groups[2].Value;
                    var titleAttrMatch = Regex.Match(attributes, @"title\s*=\s*[""']([^""']+)[""']", RegexOptions.IgnoreCase);
                    if (titleAttrMatch.Success)
                    {
                        // Add a comment to the cell with the tooltip text
                        int commentIdx = sheet.Comments.Add(rowIndex, colIndex);
                        Comment comment = sheet.Comments[commentIdx];
                        comment.Author = "Author";
                        comment.Note = titleAttrMatch.Groups[1].Value;
                    }

                    colIndex++;
                }
                rowIndex++;
            }

            // Save the workbook to an Excel file
            string excelPath = "output.xlsx";
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"HTML has been converted to Excel and saved as '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
