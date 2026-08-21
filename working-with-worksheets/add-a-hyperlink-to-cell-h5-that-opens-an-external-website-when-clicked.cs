// Title: Add an external hyperlink to cell H5 using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, accesses the first worksheet, inserts a hyperlink in cell H5 that opens https://www.example.com, sets the display text to "Visit Example Site", and saves the file as HyperlinkDemo.xlsx.
// Keywords: Aspose.Cells hyperlink C# | add external link Excel cell | set hyperlink text Aspose.Cells | Excel hyperlink programmatically | C# Aspose.Cells tutorial
// Common Searches: Aspose.Cells add hyperlink to specific cell | C# set display text for Excel hyperlink | How to link a cell to an external website with Aspose.Cells | Create clickable URL in Excel using Aspose.Cells .NET
// Developer Intent: Insert a clickable URL into cell H5 that opens an external website.
// Use Cases: Embedding reference URLs in generated financial reports. | Building marketing templates with product page links. | Automating documentation indexes that point to online guides.
// AI Prompts: Generate code to add hyperlinks to multiple cells in a worksheet with Aspose.Cells. | Show how to modify the address and display text of an existing hyperlink in a .NET workbook. | Provide an example that forces a hyperlink to open in a new browser tab using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, inserts a hyperlink in cell H5 that opens https://www.example.com, sets the display text to "Visit Example Site", and saves the file as HyperlinkDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell H5 that points to an external website
        // Parameters: cell name, rows in range, columns in range, hyperlink address
        int hyperlinkIndex = worksheet.Hyperlinks.Add("H5", 1, 1, "https://www.example.com");

        // Optionally set the text that will be displayed in the cell
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Visit Example Site";

        // Save the workbook to a file
        workbook.Save("HyperlinkDemo.xlsx");
    }
}
