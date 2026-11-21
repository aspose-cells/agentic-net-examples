// Title: How to delete rows 10‑15 from an Excel worksheet and save the result as a PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete rows 10 to 15 and export to PDF | remove specific rows from an XLSX file and convert to PDF with Aspose.Cells .NET | C# code to delete a range of rows in Excel and generate PDF output | example of using Cells.DeleteRows then saving workbook as PDF in Aspose.Cells | how to combine row deletion and PDF conversion in Aspose.Cells for .NET
// Tags: Aspose.Cells DeleteRows method C# | export Excel worksheet to PDF Aspose.Cells | remove row range Excel .NET | save modified workbook as PDF Aspose.Cells | Aspose.Cells row deletion PDF conversion

using System;
using Aspose.Cells;

// The sample loads an existing XLSX file, deletes rows 10‑15 from the first worksheet using the Cells.DeleteRows method, and then saves the updated workbook directly to a PDF file.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete rows 10 through 15 (zero‑based index: row 10 is index 9)
        // DeleteRows(startRow, totalRows)
        sheet.Cells.DeleteRows(9, 6);

        // Save the modified workbook as a PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
