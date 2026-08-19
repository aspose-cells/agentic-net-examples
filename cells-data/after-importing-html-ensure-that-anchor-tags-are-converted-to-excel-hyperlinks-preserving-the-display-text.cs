// Title: Transform HTML anchor tags into Excel hyperlinks while keeping link text – Aspose.Cells for .NET
// Description: This C# example loads an HTML file into an Aspose.Cells Workbook, scans each cell for <a> elements, extracts the href and inner text via a regular expression, replaces the cell value with plain text, and attaches a hyperlink to the same cell using Hyperlinks.Add. The workbook is then saved as XLSX.
// Keywords: Aspose.Cells | C# HTML import | Excel hyperlink | anchor tag conversion | Hyperlinks.Add | regex href extraction | load HTML to workbook | preserve link text | Aspose.Cells .NET | Excel automation
// Common Searches: how to convert HTML links to Excel hyperlinks using Aspose.Cells | preserve anchor text when importing HTML into Excel with C# | add hyperlink to cell after HTML load Aspose.Cells | Aspose.Cells replace <a> tag with hyperlink | C# regex extract href from HTML for Excel
// Developer Intent: Automatically replace HTML <a> elements in imported cells with native Excel hyperlinks that retain the displayed text.
// Use Cases: Migrating web‑based reports containing hyperlinks into Excel workbooks. | Processing HTML email newsletters to generate Excel files with functional links. | Cleaning up data after bulk HTML import so that cells show only link text but remain clickable. | Creating Excel dashboards from HTML sources where link navigation must be retained.
// AI Prompts: Generate C# code that reads an HTML file with Aspose.Cells, finds <a> tags in cell values, substitutes them with plain text, and adds matching hyperlinks to the same cells. | Explain the steps to use Aspose.Cells Hyperlinks.Add after parsing anchor tags with a regular expression. | Provide a strategy for handling multiple <a> tags inside a single cell when converting HTML to Excel using Aspose.Cells.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// This C# example loads an HTML file into an Aspose.Cells Workbook, scans each cell for <a> elements, extracts the href and inner text via a regular expression, replaces the cell value with plain text, and attaches a hyperlink to the same cell using Hyperlinks.Add. The workbook is then saved as XLSX.
class Program
{
    static void Main()
    {
        // Load the HTML file into a workbook
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Regex to capture <a href="url">display text</a>
        Regex anchorRegex = new Regex(
            @"<a\s+[^>]*href=['""]([^'""]+)['""][^>]*>(.*?)</a>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        // Iterate over all used cells
        foreach (Cell cell in cells)
        {
            string cellText = cell.StringValue;
            if (string.IsNullOrEmpty(cellText))
                continue;

            Match match = anchorRegex.Match(cellText);
            if (match.Success)
            {
                // Extract URL and display text from the anchor tag
                string url = match.Groups[1].Value;
                string displayText = match.Groups[2].Value;

                // Replace cell content with the display text
                cell.PutValue(displayText);

                // Add a hyperlink to the same cell preserving the display text
                // startCellName and endCellName are the same for a single cell
                worksheet.Hyperlinks.Add(
                    cell.Name,          // startCellName
                    cell.Name,          // endCellName
                    url,                // address
                    displayText,        // textToDisplay
                    null);              // screenTip (optional)
            }
        }

        // Save the workbook to an Excel file
        workbook.Save("output.xlsx");
    }
}
