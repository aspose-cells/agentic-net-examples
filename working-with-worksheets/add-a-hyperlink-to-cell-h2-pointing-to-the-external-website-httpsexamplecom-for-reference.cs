// Title: Add an External Hyperlink to Cell H2 with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, accesses the first worksheet, adds a hyperlink to cell H2 that points to https://example.com, sets the display text to "Example Site", and saves the file as Hyperlink_H2.xlsx. Useful for embedding reference links in generated Excel reports.
// Keywords: Aspose.Cells | C# | .NET | Excel hyperlink | cell H2 | Hyperlinks.Add | external link | save workbook
// Common Searches: Aspose.Cells add hyperlink to specific cell | C# add external link to Excel cell H2 | How to set display text for Aspose.Cells hyperlink | Save workbook after inserting hyperlink Aspose.Cells
// Developer Intent: Insert an external URL into cell H2 and define its visible text using Aspose.Cells for .NET.
// Use Cases: Include a reference URL in automated financial statements. | Provide quick access to source documentation from data export files. | Add help‑desk links to template workbooks for end‑user guidance.
// AI Prompts: Show how to add multiple hyperlinks to different cells with Aspose.Cells for .NET. | Explain how to update the address and display text of an existing hyperlink in a workbook. | Generate code that adds a hyperlink with a tooltip using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, adds a hyperlink to cell H2 that points to https://example.com, sets the display text to "Example Site", and saves the file as Hyperlink_H2.xlsx. Useful for embedding reference links in generated Excel reports.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell H2 that points to https://example.com
        worksheet.Hyperlinks.Add("H2", 1, 1, "https://example.com");

        // Set the display text for the hyperlink (optional)
        int hyperlinkIndex = worksheet.Hyperlinks.Count - 1;
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Example Site";

        // Save the workbook
        workbook.Save("Hyperlink_H2.xlsx");
    }
}
