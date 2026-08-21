// Title: Add a Mailto Email Hyperlink to Cell N3 with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a mailto hyperlink into cell N3, sets the visible text to "Send Email", and saves the file as EmailHyperlink.xlsx using the Aspose.Cells API for C#.
// Keywords: Aspose.Cells | C# email hyperlink | mailto link Excel | hyperlink cell N3 | TextToDisplay property | Hyperlinks.Add example | Excel automation .NET | save workbook with hyperlink
// Common Searches: Aspose.Cells add mailto hyperlink | C# create email link in Excel cell | set hyperlink display text Aspose.Cells | how to add clickable email address to Excel using .NET | Aspose.Cells Hyperlinks.Add usage
// Developer Intent: Insert a mailto hyperlink into cell N3 and define its display text using Aspose.Cells for .NET.
// Use Cases: Generate an invoice workbook that includes a "Contact Support" email button. | Provide a feedback link in a data‑entry template that opens the user's email client. | Create a scheduled sales report with a quick‑email link to the sales team.
// AI Prompts: Generate C# code that adds a mailto hyperlink to cell N3 with display text "Send Email" using Aspose.Cells. | Show how to retrieve and modify the address or display text of an existing hyperlink in an Aspose.Cells workbook. | Explain the steps to create and save an Excel file containing an email hyperlink with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a new workbook, inserts a mailto hyperlink into cell N3, sets the visible text to "Send Email", and saves the file as EmailHyperlink.xlsx using the Aspose.Cells API for C#.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a hyperlink to cell N3 that opens an email client
                int hyperlinkIndex = worksheet.Hyperlinks.Add("N3", 1, 1, "mailto:example@domain.com");

                // Set the display text for the hyperlink
                worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Send Email";

                // Save the workbook
                string outputPath = "EmailHyperlink.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
