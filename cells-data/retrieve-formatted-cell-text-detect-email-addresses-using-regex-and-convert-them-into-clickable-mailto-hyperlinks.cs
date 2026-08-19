// Title: Create Mailto Hyperlinks from Email Text in Excel with Aspose.Cells (C#)
// Description: This C# example scans a worksheet, detects email addresses using a regular expression, retrieves each cell's formatted HTML, and adds a mailto hyperlink via HyperlinkCollection. The email becomes the display text and the workbook is saved as EmailHyperlinks.xlsx.
// Keywords: Aspose.Cells | C# | email hyperlink | mailto link | regex email detection | GetHtmlString | HyperlinkCollection.Add | Excel automation | .NET | workbook save
// Common Searches: Aspose.Cells add mailto hyperlink C# | C# regex find email in Excel Aspose | convert email text to clickable link in Excel using Aspose.Cells | GetHtmlString cell Aspose.Cells example | create email hyperlink in worksheet programmatically
// Developer Intent: Generate clickable mailto links for email addresses found in worksheet cells.
// Use Cases: Automatically convert plain email strings in generated reports to clickable links before saving the workbook. | Build a contact‑list Excel file where each email cell opens the default mail client. | Process imported data, replace email text with hyperlinked cells, and keep the original display formatting.
// AI Prompts: Write C# code using Aspose.Cells to iterate over all used cells, detect email addresses with a regex, and add a mailto hyperlink to each cell. | Show how to use GetHtmlString to obtain a cell's formatted HTML before inserting a hyperlink in Aspose.Cells. | Explain how to handle multiple email addresses in a single cell and create separate mailto hyperlinks for each using Aspose.Cells.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsEmailHyperlinkDemo
{
    // This C# example scans a worksheet, detects email addresses using a regular expression, retrieves each cell's formatted HTML, and adds a mailto hyperlink via HyperlinkCollection. The email becomes the display text and the workbook is saved as EmailHyperlinks.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data containing email addresses
            sheet.Cells["A1"].PutValue("Contact: john.doe@example.com");
            sheet.Cells["A2"].PutValue("No email here");
            sheet.Cells["A3"].PutValue("Multiple: alice@domain.com, bob@site.org");

            // Regular expression to detect email addresses
            Regex emailRegex = new Regex(@"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}");

            // Iterate through all used cells
            foreach (Cell cell in sheet.Cells)
            {
                // Retrieve the formatted text of the cell (HTML representation)
                string formattedText = cell.GetHtmlString(false);

                // Search for email addresses in the cell's plain string value
                Match match = emailRegex.Match(cell.StringValue);
                if (match.Success)
                {
                    string email = match.Value;

                    // Add a mailto hyperlink to the cell (using HyperlinkCollection.Add)
                    int linkIndex = sheet.Hyperlinks.Add(cell.Name, 1, 1, $"mailto:{email}");
                    Hyperlink link = sheet.Hyperlinks[linkIndex];

                    // Set the display text of the hyperlink to the email address
                    link.TextToDisplay = email;

                    // Optionally, replace the cell's content with just the email (preserves display)
                    cell.PutValue(email);
                }
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("EmailHyperlinks.xlsx");
        }
    }
}
