// Title: Delete row 8 from an Excel worksheet and export the workbook to PDF using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, removes the eighth row, and saves the workbook as a PDF document. | Show the steps to use Aspose.Cells to delete a specific row in a worksheet and then perform a PDF conversion in a .NET application.
// Common Searches: asp.net aspose.cells delete row 8 from excel and convert to pdf | c# remove specific row from xlsx before pdf export using aspose.cells | how to delete a row in an Excel file with Aspose.Cells and save as PDF | aspose.cells delete rows worksheet then export to pdf c# example | remove eighth row from workbook and generate pdf with aspose.cells
// Tags: Aspose.Cells worksheet row deletion | Aspose.Cells PDF export | C# Excel row removal | Convert edited Excel to PDF with Aspose.Cells | Aspose.Cells delete rows API

using Aspose.Cells;
using System;

// The sample loads 'input.xlsx' with Aspose.Cells, accesses the first worksheet, deletes the eighth row (zero‑based index 7) using DeleteRows, and then saves the modified workbook as 'output.pdf' in PDF format.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete row 8 (zero‑based index is 7) and remove 1 row
        sheet.Cells.DeleteRows(7, 1);

        // Save the updated workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
