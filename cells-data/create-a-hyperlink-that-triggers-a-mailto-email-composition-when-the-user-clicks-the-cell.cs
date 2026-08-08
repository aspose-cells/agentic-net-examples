// Title: Add a mailto: email hyperlink to an Excel cell using Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, inserts a mailto hyperlink into cell A1, sets custom display text and a screen tip, and saves the file as MailtoHyperlink.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells mailto hyperlink C# | Excel email link Aspose | Hyperlinks.Add Aspose.Cells example | set TextToDisplay Aspose hyperlink | screen tip for Excel hyperlink | save workbook with hyperlink Aspose
// Common Searches: how to create a mailto link in Excel using Aspose.Cells | Aspose.Cells C# add email hyperlink to cell | set hyperlink tooltip Aspose.Cells | Hyperlinks.Add overload parameters Aspose | generate Excel file with clickable email addresses
// Developer Intent: Insert a clickable mailto link into a worksheet cell programmatically.
// Use Cases: Build a contact directory where each name opens the default mail client. | Automate report generation that includes direct email links for support teams. | Add explanatory tooltips to email links for better user guidance.
// AI Prompts: Show how to add several mailto hyperlinks to different cells with Aspose.Cells. | Demonstrate customizing the display text and tooltip of a hyperlink in a workbook. | Provide code to enumerate existing hyperlinks and replace them with mailto URLs.

using System;
using Aspose.Cells;

// This example creates a new workbook, inserts a mailto hyperlink into cell A1, sets custom display text and a screen tip, and saves the file as MailtoHyperlink.xlsx with Aspose.Cells for .NET.
class MailtoHyperlinkDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Mailto address that will open the default email client
        string mailtoAddress = "mailto:john.doe@example.com";

        // Add a hyperlink to cell A1 using the (string, int, int, string) overload
        int hyperlinkIndex = worksheet.Hyperlinks.Add("A1", 1, 1, mailtoAddress);

        // Set the text that will be displayed in the cell
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Email John Doe";

        // Optional: add a screen tip for the hyperlink
        worksheet.Hyperlinks[hyperlinkIndex].ScreenTip = "Click to compose an email";

        // Save the workbook to an XLSX file
        workbook.Save("MailtoHyperlink.xlsx");
    }
}
