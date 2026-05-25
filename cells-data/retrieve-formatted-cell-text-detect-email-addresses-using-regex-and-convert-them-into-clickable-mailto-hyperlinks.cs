using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Net;
using Aspose.Cells;

class EmailHyperlinkDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample cells containing email addresses
        cells["A1"].PutValue("Contact: john.doe@example.com for more info.");
        cells["B2"].PutValue("Team: alice@domain.com, bob@site.org");
        cells["C3"].PutValue("No email address here.");

        // Regular expression to detect email addresses (case‑insensitive)
        Regex emailRegex = new Regex(@"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", RegexOptions.IgnoreCase);

        // Iterate through all used cells
        foreach (Cell cell in cells)
        {
            string plainText = cell.StringValue;
            if (string.IsNullOrEmpty(plainText))
                continue;

            // Find all email matches in the cell text
            MatchCollection matches = emailRegex.Matches(plainText);
            if (matches.Count == 0)
                continue;

            // Build an HTML string where each email is replaced with a mailto hyperlink
            StringBuilder htmlBuilder = new StringBuilder();
            int lastPos = 0;

            foreach (Match match in matches)
            {
                // Append text before the email (HTML‑encoded)
                htmlBuilder.Append(WebUtility.HtmlEncode(plainText.Substring(lastPos, match.Index - lastPos)));

                // Append the mailto hyperlink for the email
                string email = match.Value;
                htmlBuilder.Append($"<a href=\"mailto:{email}\">{email}</a>");

                // Update the last processed position
                lastPos = match.Index + match.Length;
            }

            // Append any remaining text after the last email (HTML‑encoded)
            htmlBuilder.Append(WebUtility.HtmlEncode(plainText.Substring(lastPos)));

            // Set the cell's HTML string so the hyperlink becomes clickable
            cell.HtmlString = htmlBuilder.ToString();
        }

        // Save the workbook with the added mailto hyperlinks
        workbook.Save("EmailHyperlinks.xlsx");
    }
}