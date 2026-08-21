// Title: Add Clickable Hyperlinks from URLs in Cell Text with Aspose.Cells for .NET
// Description: Creates a workbook, inserts a string that may contain an HTTP/HTTPS URL, reads the cell's formatted text, uses a regular expression to find a URL, and adds a hyperlink to the same cell while preserving the original display text before saving the file.
// Keywords: Aspose.Cells hyperlink C# | detect URL in Excel cell | regex URL detection Aspose.Cells | add hyperlink programmatically .NET | cell StringValue Aspose.Cells | preserve cell formatting hyperlink
// Common Searches: Aspose.Cells add hyperlink from cell text | detect URLs in Excel cells using C# | convert cell string to clickable link Aspose.Cells | regex find URL in worksheet cell | preserve formatting when adding hyperlink Aspose.Cells
// Developer Intent: Find URLs inside a cell's formatted value and turn them into clickable hyperlinks within the same workbook.
// Use Cases: Automatically convert any URL embedded in a cell to an active hyperlink while keeping the original sentence as the display text. | Batch‑process a worksheet to locate URLs in each cell and insert hyperlinks before exporting the workbook. | Generate reports where web addresses become clickable links without altering existing cell styles.
// AI Prompts: Write C# code with Aspose.Cells that scans all cells in a worksheet, detects URLs using a regular expression, and adds hyperlinks that display the original cell text. | Show how to use Hyperlinks.Add to create a hyperlink in the same cell after extracting a URL from the cell's StringValue. | Explain how to preserve cell styling while converting detected URLs into clickable links in an Aspose.Cells workbook.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHyperlinkDemo
{
    // Creates a workbook, inserts a string that may contain an HTTP/HTTPS URL, reads the cell's formatted text, uses a regular expression to find a URL, and adds a hyperlink to the same cell while preserving the original display text before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example: put a string that may contain a URL into cell B2
            Cell targetCell = sheet.Cells["B2"];
            targetCell.PutValue("Visit https://www.example.com for more info");

            // Retrieve the formatted string value of the cell
            string formattedText = targetCell.StringValue; // formatted according to cell style

            // Simple URL detection using regular expression
            // This pattern matches http or https URLs
            string urlPattern = @"https?://[^\s]+";
            Match match = Regex.Match(formattedText, urlPattern, RegexOptions.IgnoreCase);

            if (match.Success)
            {
                string url = match.Value;

                // Add a hyperlink to the same cell (lifecycle rule: use Hyperlinks.Add)
                // The Add method returns the index of the newly created hyperlink
                int hyperlinkIndex = sheet.Hyperlinks.Add(targetCell.Name, 1, 1, url);

                // Optionally set the display text of the hyperlink to the original cell text
                Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
                hyperlink.TextToDisplay = formattedText;

                Console.WriteLine($"Hyperlink added to {targetCell.Name}: {url}");
            }
            else
            {
                Console.WriteLine("No URL detected in the cell.");
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("OutputWithHyperlink.xlsx", SaveFormat.Xlsx);
        }
    }
}
