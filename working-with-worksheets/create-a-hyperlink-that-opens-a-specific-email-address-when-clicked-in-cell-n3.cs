// Title: Aspose.Cells for .NET – Add a Mailto Email Hyperlink to Cell N3 (C#)
// Description: C# example that creates a new Workbook with Aspose.Cells, accesses the first worksheet, inserts a "mailto:" hyperlink into cell N3, sets custom display text, and saves the file as EmailHyperlink.xlsx. Ideal for programmatically adding email links in Excel reports.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# | mailto hyperlink | Excel email link | cell N3 | hyperlink API | save workbook | GitHub example | email hyperlink C#
// Common Searches: Aspose.Cells add mailto link C# | Create email hyperlink in Excel using Aspose.Cells | How to insert a mailto hyperlink into a specific cell with Aspose.Cells | Set hyperlink display text Aspose.Cells C# | Save workbook with hyperlink Aspose.Cells
// Developer Intent: Insert an email (mailto) hyperlink into cell N3 of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Add a one‑click "Send Email" button in generated invoices for quick client contact. | Provide a feedback email link in automated reports or dashboards. | Create a contact cell in a data‑entry template that opens the user's default mail client.
// AI Prompts: Generate C# code with Aspose.Cells that adds a mailto hyperlink to cell N3 and customizes the display text. | Explain how to update the hyperlink target and display text for multiple cells using Aspose.Cells. | Show error handling for adding a hyperlink when the target cell already contains a link.

using System;
using Aspose.Cells;

// C# example that creates a new Workbook with Aspose.Cells, accesses the first worksheet, inserts a "mailto:" hyperlink into cell N3, sets custom display text, and saves the file as EmailHyperlink.xlsx. Ideal for programmatically adding email links in Excel reports.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell N3 that opens an email client
        // The address uses the "mailto:" scheme
        int hyperlinkIndex = worksheet.Hyperlinks.Add("N3", 1, 1, "mailto:someone@example.com");

        // Optional: set the text that will be displayed in the cell
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Send Email";

        // Save the workbook to a file
        workbook.Save("EmailHyperlink.xlsx");
    }
}
