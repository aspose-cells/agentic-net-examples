// Title: C# Aspose.Cells example: delete rows 60‑65 from an Excel worksheet and save as PDF
// Common Searches: Aspose.Cells C# delete rows 60 to 65 from Excel worksheet | How to remove a range of rows in Excel using Aspose.Cells and export to PDF | C# code to delete specific rows in an Excel file and convert to PDF with Aspose.Cells | Export modified Excel workbook to PDF after deleting rows with Aspose.Cells .NET
// Tags: Aspose.Cells DeleteRows method usage | Excel to PDF conversion after row deletion | C# delete rows from worksheet | Aspose.Cells save workbook as PDF | Batch row removal in Excel with Aspose.Cells

using Aspose.Cells;
using System;

// The sample loads 'input.xlsx' with Aspose.Cells, deletes rows 60‑65 (zero‑based index 59, count 6) from the first worksheet, and saves the workbook as 'output.pdf' in PDF format.
class Program
{
    static void Main()
    {
        // Load the Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete rows 60 through 65.
        // Aspose.Cells uses zero‑based indexing, so row 60 is index 59.
        // Delete 6 rows starting at index 59.
        worksheet.Cells.DeleteRows(59, 6);

        // Save the result as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
