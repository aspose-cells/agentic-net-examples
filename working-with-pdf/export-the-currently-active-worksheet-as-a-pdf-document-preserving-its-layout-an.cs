using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExportActiveWorksheetToPdf
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Data 2");
            sheet.Cells["B3"].PutValue(456);

            // Optionally set a different worksheet as active
            // workbook.Worksheets.ActiveSheetIndex = 0; // default is first sheet

            // Configure PDF save options to export only the active worksheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.SheetSet = SheetSet.Active; // export only active sheet
            // Preserve layout and formatting (default behavior)

            // Save the active worksheet as a PDF file
            workbook.Save("ActiveWorksheet.pdf", pdfOptions); // save
        }
    }
}