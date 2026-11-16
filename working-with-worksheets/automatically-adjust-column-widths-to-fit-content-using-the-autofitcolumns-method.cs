// Title: Automatically adjust Excel column widths to fit cell content with Aspose.Cells AutoFitColumns in C#
// Common Searches: C# Aspose.Cells AutoFitColumns to automatically size worksheet columns | How to resize Excel columns to fit content using Aspose.Cells .NET API | Example of adjusting column width to cell data with Aspose.Cells in C# | AutoFitColumns method usage for fitting column widths in an Excel workbook | Programmatically auto‑fit all columns after populating data with Aspose.Cells
// Tags: AutoFitColumns method Aspose.Cells C# | auto resize Excel column width programmatically | fit worksheet columns to content Aspose.Cells | adjust column width using Aspose.Cells .NET | Excel column auto‑fit example C#

using System;
using Aspose.Cells;

// The sample creates a new Workbook, writes headers and data to cells, calls sheet.AutoFitColumns() to automatically size every column to its content, and saves the file as AutoFitColumnsExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data.
        sheet.Cells["A1"].PutValue("Header 1");
        sheet.Cells["B1"].PutValue("Header 2");
        sheet.Cells["A2"].PutValue("Short");
        sheet.Cells["B2"].PutValue("This is a longer piece of text that needs more space");

        // Automatically adjust all columns in the worksheet to fit the content.
        sheet.AutoFitColumns();

        // Save the workbook to a file.
        workbook.Save("AutoFitColumnsExample.xlsx");
    }
}
