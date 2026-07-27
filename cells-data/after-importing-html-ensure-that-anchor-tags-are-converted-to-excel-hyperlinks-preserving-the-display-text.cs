using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the HTML file. HtmlLoadOptions can be customized if needed.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        // Example: treat any text starting with '=' as plain text, not a formula.
        loadOptions.HasFormula = false;

        // Replace "input.html" with the path to your HTML source.
        Workbook workbook = new Workbook("input.html", loadOptions);
        Worksheet sheet = workbook.Worksheets[0];

        // Regular expression to match an anchor tag and capture href and inner text.
        // Supports both single and double quotes.
        Regex anchorRegex = new Regex(@"<a\s+[^>]*href\s*=\s*['""](?<url>[^'""]+)['""][^>]*>(?<text>.*?)</a>", RegexOptions.IgnoreCase);

        // Iterate through all used cells.
        foreach (Cell cell in sheet.Cells)
        {
            string cellValue = cell.StringValue;
            if (string.IsNullOrEmpty(cellValue))
                continue;

            Match match = anchorRegex.Match(cellValue);
            if (match.Success)
            {
                // Extract URL and display text.
                string url = match.Groups["url"].Value;
                string displayText = match.Groups["text"].Value;

                // Replace the cell content with the display text.
                cell.PutValue(displayText);

                // Add a hyperlink to the cell.
                int hyperlinkIndex = sheet.Hyperlinks.Add(cell.Name, 1, 1, url);
                // Set the text to display (optional, as we already set cell value).
                sheet.Hyperlinks[hyperlinkIndex].TextToDisplay = displayText;
            }
        }

        // Save the workbook to Excel format (or any other supported format).
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}