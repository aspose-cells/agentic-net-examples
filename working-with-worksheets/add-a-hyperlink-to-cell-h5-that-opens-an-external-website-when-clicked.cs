// Title: Add an External Hyperlink to Cell H5 with Display Text using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts a hyperlink in cell H5 that opens https://www.example.com, sets the visible text to "Visit Example Site", and saves the file as HyperlinkDemo.xlsx.
// Keywords: Aspose.Cells | C# hyperlink | add hyperlink to Excel cell | external URL | cell H5 | display text | save workbook | Excel automation
// Common Searches: Aspose.Cells add hyperlink to cell | C# set hyperlink display text in Excel | Create clickable link in Excel using Aspose.Cells | Programmatically add external URL to Excel cell
// Developer Intent: Insert an external URL into cell H5 and customize the text shown in the cell.
// Use Cases: Generate a report where cell H5 links to a live dashboard or documentation site. | Build an Excel‑based navigation sheet that directs users to online resources. | Automate the addition of labeled web links across multiple worksheets for quick access.
// AI Prompts: Write C# code with Aspose.Cells to add a hyperlink to cell H5 that opens https://www.example.com and displays "Visit Example Site". | Show how to add multiple external hyperlinks with custom display text to different cells using Aspose.Cells for .NET. | Explain how to update the address or display text of an existing hyperlink in an Aspose.Cells worksheet.

using System;
using Aspose.Cells;

// Creates a new workbook, inserts a hyperlink in cell H5 that opens https://www.example.com, sets the visible text to "Visit Example Site", and saves the file as HyperlinkDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell H5 (row 5, column H) that points to an external website
        // The Add method returns the index of the newly created hyperlink
        int hyperlinkIndex = worksheet.Hyperlinks.Add("H5", 1, 1, "https://www.example.com");

        // Optionally set the text that will be displayed in the cell
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Visit Example Site";

        // Save the workbook to a file
        workbook.Save("HyperlinkDemo.xlsx");
    }
}
