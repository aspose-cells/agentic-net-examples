// Title: Insert a clickable external URL hyperlink into cell B2 of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to place the text "Google" in cell B2, attach a hyperlink to https://www.google.com that opens in a new browser tab, and save the workbook as HyperlinkExample.xlsx. | Show how to call the Hyperlinks.Add method in Aspose.Cells to add an external URL to a specific worksheet cell and configure it to launch in a new tab within a .NET application.
// Common Searches: Aspose.Cells add external URL hyperlink to Excel cell C# | C# Aspose.Cells Hyperlinks.Add open link in new tab | how to make Excel hyperlink open in new window using Aspose.Cells | example of inserting clickable link into specific cell with Aspose.Cells for .NET
// Tags: Aspose.Cells Hyperlinks.Add method | external URL hyperlink in Excel worksheet C# | hyperlink opens new tab Aspose.Cells | generate workbook with clickable link using Aspose.Cells

using Aspose.Cells;
using System;

// // Creates a new workbook, writes "Google" into B2, adds a URL hyperlink to https://www.google.com that opens in a new tab, and saves the file as HyperlinkExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put display text into the target cell (B2)
            Cell cell = sheet.Cells["B2"];
            cell.PutValue("Google");

            // Add a hyperlink to the cell that points to an external website
            // The overload without HyperlinkType defaults to a URL hyperlink
            sheet.Hyperlinks.Add(cell.Row, cell.Column, 1, 1, "https://www.google.com");

            // Save the workbook to a file
            workbook.Save("HyperlinkExample.xlsx");
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
