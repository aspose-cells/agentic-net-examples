// Title: C# Aspose.Cells example: Detect URLs in a cell’s displayed text and add a clickable hyperlink
// AI Prompts: Generate C# code using Aspose.Cells that reads a cell’s DisplayStringValue, extracts any http/https URL with a regular expression, and inserts a hyperlink into the same cell while preserving the original sentence as the link text. | Create a reusable C# method for Aspose.Cells that accepts a Worksheet and cell address, finds the first URL in the cell’s formatted string, adds a hyperlink pointing to that URL, and returns the hyperlink index. | Show how to modify an existing Aspose.Cells workbook so that cells containing URLs in their displayed text are automatically converted to clickable hyperlinks during a save operation.
// Common Searches: Aspose.Cells C# add hyperlink to a cell based on its displayed string | How to extract URL from Excel cell text using Aspose.Cells .NET | C# regex to find http links in a worksheet cell with Aspose.Cells | Convert plain text URL in Excel to clickable link using Aspose.Cells API | Programmatically create hyperlink in .xlsx file from cell content in C#
// Tags: Aspose.Cells add hyperlink from cell text | DisplayStringValue URL extraction Aspose.Cells | C# regex hyperlink creation Excel | Aspose.Cells .xlsx hyperlink insertion | programmatic Excel hyperlink generation C#

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHyperlinkDemo
{
    // Demonstrates retrieving a cell's displayed string with DisplayStringValue, using a regular expression to locate an http/https URL, adding a hyperlink to the same cell while keeping the original text as the link label, and saving the workbook as an .xlsx file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample cell that may contain a URL
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Visit our site at https://www.example.com for more info.");

            // Retrieve the formatted string value of the cell
            // DisplayStringValue returns the value as shown in Excel (respecting cell style)
            string formattedText = cell.DisplayStringValue;

            // Simple regex to detect URLs starting with http or https
            Regex urlRegex = new Regex(@"(http|https)://[^\s]+", RegexOptions.IgnoreCase);
            Match match = urlRegex.Match(formattedText);

            if (match.Success)
            {
                string url = match.Value;

                // Add a hyperlink to the same cell (A1) pointing to the detected URL
                // The Add method returns the index of the created hyperlink
                int hyperlinkIndex = sheet.Hyperlinks.Add("A1", 1, 1, url);

                // Optionally set the display text of the hyperlink to the original cell text
                Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
                hyperlink.TextToDisplay = formattedText;
            }

            // Save the workbook to a file
            workbook.Save("HyperlinkResult.xlsx", SaveFormat.Xlsx);
        }
    }
}
