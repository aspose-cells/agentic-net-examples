using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PrintTitleRowsPdfDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a header row that will be repeated on each printed page
        sheet.Cells["A1"].PutValue("Header A");
        sheet.Cells["B1"].PutValue("Header B");
        sheet.Cells["C1"].PutValue("Header C");

        // Populate sample data rows
        for (int i = 2; i <= 50; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Data {i - 1} - A");
            sheet.Cells[$"B{i}"].PutValue($"Data {i - 1} - B");
            sheet.Cells[$"C{i}"].PutValue($"Data {i - 1} - C");
        }

        // Set the rows to repeat at the top of each PDF page (first row)
        sheet.PageSetup.PrintTitleRows = "$1:$1";

        // Define the print area to include all populated cells
        sheet.PageSetup.PrintArea = "A1:C50";

        // Create PDF save options (default settings are sufficient)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file
        workbook.Save("PrintTitleRowsDemo.pdf", pdfOptions);
    }
}