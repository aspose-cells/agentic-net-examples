// Title: Delete rows 10‑15 from an Excel worksheet and export the cleaned sheet to PDF with Aspose.Cells in C#
// AI Prompts: Load an .xlsx workbook, delete a six‑row block beginning at row 10, and save the result as a PDF with Aspose.Cells in C#. | Write C# code that calls Cells.DeleteRows to eliminate a specific row range and then exports the worksheet to PDF. | Create a script that opens an Excel file, removes a contiguous set of rows from the first sheet, and converts the modified workbook to PDF using Aspose.Cells.
// Common Searches: Aspose.Cells C# delete rows 10 to 15 then save as PDF | How to remove a specific row range from an Excel file before PDF conversion using Aspose.Cells | C# example for deleting rows in a worksheet and exporting to PDF with Aspose.Cells | Delete rows 9‑14 zero based Aspose.Cells and generate PDF output
// Tags: Aspose.Cells DeleteRows method C# | Excel worksheet row removal Aspose.Cells | SaveFormat.Pdf export Aspose.Cells | C# workbook modification and PDF generation | Aspose.Cells row range deletion example

using Aspose.Cells;

// The C# program loads 'input.xlsx' with Aspose.Cells, removes rows 10‑15 (zero‑based index 9, total 6 rows) from the first worksheet, and saves the updated workbook as 'output.pdf' using the SaveFormat.Pdf option.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        var workbook = new Workbook("input.xlsx");

        // Get the first worksheet (index 0)
        var worksheet = workbook.Worksheets[0];

        // Delete rows 10 through 15.
        // Aspose.Cells uses zero‑based indexing, so row 10 is index 9.
        // DeleteRows(startRow, totalRows) where totalRows = 6 (rows 10‑15 inclusive).
        worksheet.Cells.DeleteRows(9, 6);

        // Save the cleaned worksheet as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
