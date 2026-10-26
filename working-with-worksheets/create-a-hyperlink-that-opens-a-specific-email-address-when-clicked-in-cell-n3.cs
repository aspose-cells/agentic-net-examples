// Title: Create a mailto hyperlink in cell N3 of an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to add a mailto link to a specific cell with Aspose.Cells C# | Aspose.Cells C# insert email hyperlink into Excel cell N3 | C# create clickable email address in Excel using Aspose.Cells | add email hyperlink to worksheet cell using Aspose.Cells .NET API | set mailto hyperlink for cell N3 in Aspose.Cells workbook
// Tags: Aspose.Cells mailto hyperlink insertion | C# worksheet cell email link | Excel cell N3 clickable email using Aspose | generate email hyperlink in .xlsx with Aspose.Cells | save workbook containing email hyperlink

using Aspose.Cells;
using System;

// The example creates a new workbook, writes display text to cell N3, adds a mailto hyperlink that opens an email to john.doe@example.com, and saves the file as HyperlinkEmail.xlsx.
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

            // Set the display text in cell N3
            Cell cell = sheet.Cells["N3"];
            cell.PutValue("Email John Doe");

            // Add a mailto hyperlink to the same cell
            // Parameters: row, column, height (rows), width (columns), address
            sheet.Hyperlinks.Add(cell.Row, cell.Column, 1, 1, "mailto:john.doe@example.com");

            // Save the workbook
            workbook.Save("HyperlinkEmail.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
