using System;
using Aspose.Cells;

class MailtoHyperlinkDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Target cell for the hyperlink
        string cellName = "A1";

        // Mailto address that opens the default email client
        string mailtoAddress = "mailto:john.doe@example.com";

        // Add the hyperlink to the specified cell
        int hyperlinkIndex = worksheet.Hyperlinks.Add(cellName, 1, 1, mailtoAddress);

        // Set custom display text for the hyperlink
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Email John Doe";

        // Save the workbook to a file
        workbook.Save("MailtoHyperlink.xlsx");
    }
}