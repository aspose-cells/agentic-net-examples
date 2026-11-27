// Title: Add a mailto hyperlink to cell N3 in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add email hyperlink to a specific cell | Create a mailto link in Excel cell N3 using .NET | How to insert a clickable email address in an Aspose.Cells worksheet | C# code to set a mailto hyperlink on a single cell in an XLSX file | Aspose.Cells example for adding a mailto hyperlink to a workbook
// Tags: Aspose.Cells create mailto link | C# insert hyperlink into Excel cell | Excel N3 hyperlink Aspose.Cells | Aspose.Cells save workbook as xlsx | define 1x1 hyperlink range Aspose

using Aspose.Cells;
using System;

// The sample creates a new workbook, writes "Send Email" into cell N3, adds a mailto hyperlink that opens support@example.com covering a 1 × 1 range, and saves the file as HyperlinkEmail.xlsx using Aspose.Cells for .NET.
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

            // Target cell N3 (zero‑based indexing)
            Cell cell = sheet.Cells["N3"];
            cell.PutValue("Send Email"); // Display text for the hyperlink

            // Add a mailto hyperlink to the cell (1 row × 1 column span)
            sheet.Hyperlinks.Add(cell.Row, cell.Column, 1, 1, "mailto:support@example.com");

            // Save the workbook
            workbook.Save("HyperlinkEmail.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
