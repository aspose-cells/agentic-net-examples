// Title: Add a hyperlink to an external PDF that opens in a new browser tab using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add hyperlink to PDF that opens in new tab | How to create a cell link to an external PDF with Aspose.Cells .NET | C# Aspose.Cells set hyperlink target attribute for external PDF | Saving Excel workbook with PDF hyperlink using Aspose.Cells | Open PDF from Excel cell in browser using Aspose.Cells C#
// Tags: Hyperlinks.Add external PDF Aspose.Cells | C# Excel cell hyperlink new tab | Aspose.Cells workbook save with hyperlink | Open PDF from Excel cell using Aspose.Cells | Aspose.Cells hyperlink target attribute

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, writes "Open PDF" into cell B2, adds a hyperlink to an external PDF (http://example.com/document.pdf) that opens in a new browser tab, and saves the file as HyperlinkExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set display text for the hyperlink
            Cell cell = sheet.Cells["B2"];
            cell.PutValue("Open PDF");

            // Add a hyperlink to the specified cell (B2)
            // Parameters: cell name, first row, first column (unused here), hyperlink address
            sheet.Hyperlinks.Add("B2", 0, 0, "http://example.com/document.pdf");

            // Define output file path
            string outputPath = "HyperlinkExample.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
