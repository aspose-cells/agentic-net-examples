// Title: Create and style a mailto: hyperlink in an Excel cell with Aspose.Cells for .NET (C#)
// AI Prompts: Insert a mailto: hyperlink into cell A1, set its display text to "Send Email", and add a screen tip using Aspose.Cells in C#. | Apply blue font color and single underline to the hyperlink cell so it appears as a typical web link. | Save the workbook as an XLSX file named MailtoHyperlinkDemo.xlsx after configuring the email link.
// Common Searches: Aspose.Cells C# add mailto hyperlink to Excel cell | how to set hyperlink display text and screen tip with Aspose.Cells | format Excel cell as blue underlined hyperlink using Aspose.Cells .NET | save workbook with email link using Aspose.Cells C# example | Aspose.Cells create clickable email link in worksheet
// Tags: Aspose.Cells add mailto hyperlink | C# Aspose.Cells set hyperlink display text | Aspose.Cells hyperlink screen tip | format cell as blue underlined link Aspose.Cells | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkMailtoDemo
{
    // The example creates a new workbook, adds a mailto: hyperlink to cell A1 with custom display text and a screen tip, styles the cell with blue underlined font to look like a web link, and saves the file as MailtoHyperlinkDemo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell where the hyperlink will be placed
            string targetCell = "A1";

            // Add a hyperlink to the cell using the Add method (cell name, rows, columns, address)
            // The address uses the "mailto:" scheme to open the default email client
            int hyperlinkIndex = worksheet.Hyperlinks.Add(targetCell, 1, 1, "mailto:someone@example.com");

            // Retrieve the created hyperlink to customize its display text and screen tip
            Hyperlink mailtoLink = worksheet.Hyperlinks[hyperlinkIndex];
            mailtoLink.TextToDisplay = "Send Email";
            mailtoLink.ScreenTip = "Click to compose an email";

            // Optionally, style the cell to look like a typical hyperlink (blue and underlined)
            Style linkStyle = worksheet.Cells[targetCell].GetStyle();
            linkStyle.Font.Color = System.Drawing.Color.Blue;
            linkStyle.Font.Underline = FontUnderlineType.Single;
            worksheet.Cells[targetCell].SetStyle(linkStyle);

            // Save the workbook to an XLSX file
            workbook.Save("MailtoHyperlinkDemo.xlsx");
        }
    }
}
