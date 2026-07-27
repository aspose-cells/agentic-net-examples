using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHyperlinkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Example: put a string that may contain a URL into cell A1
            Cell targetCell = sheet.Cells["A1"];
            targetCell.PutValue("Check out https://www.example.com for more info.");

            // Retrieve the formatted string of the cell (as displayed in Excel)
            string formattedText = targetCell.DisplayStringValue;

            // Simple URL detection using regular expression
            // This pattern matches http or https URLs until a whitespace or end of line
            Regex urlRegex = new Regex(@"https?://\S+", RegexOptions.IgnoreCase);
            Match match = urlRegex.Match(formattedText);

            if (match.Success && !targetCell.ContainsExternalLink)
            {
                string url = match.Value;

                // Add a hyperlink to the same cell using the Hyperlinks collection
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, address
                int hyperlinkIndex = sheet.Hyperlinks.Add(targetCell.Row, targetCell.Column, 1, 1, url);

                // Optionally set the display text of the hyperlink to the original cell text
                Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
                hyperlink.TextToDisplay = formattedText;

                // If you want the cell to show only the display text (without the raw URL),
                // you can overwrite the cell value with the display text
                targetCell.PutValue(formattedText);
            }

            // Save the workbook to verify the result
            workbook.Save("HyperlinkResult.xlsx", SaveFormat.Xlsx);
        }
    }
}