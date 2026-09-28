// Title: Insert a mailto hyperlink into cell N3 with display text using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to add a mailto hyperlink to cell N3 and set the visible text to "Contact Support". | Show how to create a clickable email link in a specific Excel cell with Aspose.Cells, including saving the workbook.
// Common Searches: asp.net add mailto link to Excel cell using Aspose.Cells | c# aspose.cells set hyperlink display text for a cell | how to create email hyperlink in a specific worksheet cell with Aspose.Cells | aspose.cells insert clickable email address into N3 cell
// Tags: Aspose.Cells mailto hyperlink | C# set hyperlink text Aspose.Cells | Excel cell email link Aspose.Cells | Hyperlink range creation Aspose.Cells | Save workbook with email hyperlink Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, adds a mailto hyperlink to cell N3 with the display text "Contact Support", and saves the file as HyperlinkEmail.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the email address (mailto link)
        string emailLink = "mailto:support@example.com";

        // Add a hyperlink to cell N3 (row index 2, column index 13)
        // This creates a 1x1 hyperlink range that opens the default mail client
        sheet.Hyperlinks.Add(2, 13, 1, 1, emailLink);

        // Set the display text for the hyperlink
        sheet.Cells["N3"].PutValue("Contact Support");

        // Save the workbook
        workbook.Save("HyperlinkEmail.xlsx");
    }
}
