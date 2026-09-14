// Title: Insert a hyperlink with custom display text and screen tip into a specific Excel cell using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to add a hyperlink with the text "Visit Aspose" and a screen tip to cell A1, then save the workbook. | Show how to call worksheet.Hyperlinks.Add to set the address, display text, and screen tip for an Excel cell in a .NET application. | Generate a complete example that creates a workbook, inserts a hyperlink with custom label and tooltip, and writes the file to disk using Aspose.Cells.
// Common Searches: aspnet add hyperlink with display text to Excel cell using Aspose.Cells | how to set screen tip for a hyperlink in Aspose.Cells C# | example of HyperlinkCollection.Add method for Excel workbook .NET | save Excel file after inserting hyperlink with Aspose.Cells library | C# code to create hyperlink in cell A1 with custom text using Aspose.Cells
// Tags: Aspose.Cells HyperlinkCollection.Add | insert hyperlink into Excel cell C# | custom hyperlink display text Aspose.Cells | hyperlink screen tip .NET | save workbook with Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, accesses the first worksheet, adds a hyperlink to cell A1 with the display text "Visit Aspose" and a screen tip, then saves the workbook as HyperlinkDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell A1 with display text and a URL
        // Parameters: startCell, endCell, address, textToDisplay, screenTip
        worksheet.Hyperlinks.Add("A1", "A1", "https://www.aspose.com", "Visit Aspose", "Open Aspose website");

        // Save the workbook to a file
        workbook.Save("HyperlinkDemo.xlsx");
    }
}
