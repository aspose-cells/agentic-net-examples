// Title: Add a mailto: email hyperlink to a specific Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, writes display text to a cell, attaches an email link with Aspose.Cells, and saves the file as XLSX. | Provide a .NET snippet that uses Aspose.Cells Hyperlinks.Add to embed an email address in cell A1 and export the workbook.
// Common Searches: asp.net add mailto hyperlink to Excel cell with Aspose.Cells | c# Aspose.Cells example for inserting email link into worksheet | how to use Hyperlinks.Add to create a clickable email address in generated XLSX | programmatically generate Excel file with email hyperlink using Aspose.Cells .NET
// Tags: Aspose.Cells add mailto hyperlink | C# Hyperlinks.Add usage in Excel workbook | email hyperlink insertion in XLSX with .NET | programmatic Excel email hyperlink Aspose.Cells | save workbook with mailto link using Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new workbook, writes "Send Email" into cell A1, adds a mailto: hyperlink to that cell via the Hyperlinks.Add method, and saves the workbook as HyperlinkMailto.xlsx, with exception handling for any runtime errors.
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

            // Put display text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Send Email");

            // Add a mailto hyperlink to cell A1
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
            sheet.Hyperlinks.Add(0, 0, 0, 0, "mailto:someone@example.com");

            // Save the workbook
            workbook.Save("HyperlinkMailto.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
