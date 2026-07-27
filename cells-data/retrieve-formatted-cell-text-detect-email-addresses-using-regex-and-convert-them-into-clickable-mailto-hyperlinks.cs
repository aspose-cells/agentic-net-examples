using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsEmailHyperlinkDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data containing emails in different contexts
            cells["A1"].PutValue("Contact: john.doe@example.com for details.");
            cells["A2"].PutValue("support@mydomain.org");
            cells["A3"].PutValue("No email here");
            cells["A4"].PutValue("Multiple emails: alice@mail.com, bob@mail.com");

            // Regular expression to detect email addresses
            Regex emailRegex = new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", RegexOptions.Compiled);

            // Iterate through all used cells
            foreach (Cell cell in cells)
            {
                // Get the plain string value of the cell
                string cellText = cell.StringValue;
                if (string.IsNullOrEmpty(cellText))
                    continue;

                // Find all email matches in the cell text
                MatchCollection matches = emailRegex.Matches(cellText);
                if (matches.Count == 0)
                    continue;

                // If the cell contains only a single email, add a standard hyperlink
                if (matches.Count == 1 && cellText.Trim().Equals(matches[0].Value, StringComparison.OrdinalIgnoreCase))
                {
                    // Add a mailto hyperlink to the whole cell
                    int index = sheet.Hyperlinks.Add(cell.Name, 1, 1, "mailto:" + matches[0].Value);
                    // Set the displayed text to the email address
                    sheet.Hyperlinks[index].TextToDisplay = matches[0].Value;
                }
                else
                {
                    // For cells with mixed content, replace each email with an HTML mailto link
                    string html = cellText;
                    foreach (Match m in matches)
                    {
                        string email = m.Value;
                        string mailtoLink = $"<a href=\"mailto:{email}\">{email}</a>";
                        // Replace the plain email with the HTML link (case‑insensitive)
                        html = Regex.Replace(html, Regex.Escape(email), mailtoLink, RegexOptions.IgnoreCase);
                    }

                    // Set the cell's HtmlString so the hyperlink becomes clickable
                    cell.HtmlString = html;
                }
            }

            // Save the workbook
            workbook.Save("EmailHyperlinks.xlsx");
        }
    }
}