using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfReport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare worksheets with sample data
            // -------------------------------------------------
            // Summary sheet (first sheet)
            Worksheet summarySheet = workbook.Worksheets[0];
            summarySheet.Name = "Summary";

            // Add two data worksheets
            int sheetIndex1 = workbook.Worksheets.Add();
            Worksheet dataSheet1 = workbook.Worksheets[sheetIndex1];
            dataSheet1.Name = "Sales_Q1";
            dataSheet1.Cells["A1"].PutValue("Product");
            dataSheet1.Cells["B1"].PutValue("Revenue");
            dataSheet1.Cells["A2"].PutValue("Widget");
            dataSheet1.Cells["B2"].PutValue(12000);

            int sheetIndex2 = workbook.Worksheets.Add();
            Worksheet dataSheet2 = workbook.Worksheets[sheetIndex2];
            dataSheet2.Name = "Sales_Q2";
            dataSheet2.Cells["A1"].PutValue("Product");
            dataSheet2.Cells["B1"].PutValue("Revenue");
            dataSheet2.Cells["A2"].PutValue("Gadget");
            dataSheet2.Cells["B2"].PutValue(15000);

            // -------------------------------------------------
            // 2. Add hyperlinks on the summary sheet that
            //    open the corresponding worksheets when clicked
            // -------------------------------------------------
            // Header
            summarySheet.Cells["A1"].PutValue("Worksheet");
            summarySheet.Cells["B1"].PutValue("Link");

            // Hyperlink to Sales_Q1
            summarySheet.Cells["A2"].PutValue(dataSheet1.Name);
            // Internal link format: 'SheetName'!CellAddress
            string linkToSheet1 = $"'{dataSheet1.Name}'!A1";
            summarySheet.Hyperlinks.Add(1, 1, 1, 1, linkToSheet1); // B2 cell

            // Hyperlink to Sales_Q2
            summarySheet.Cells["A3"].PutValue(dataSheet2.Name);
            string linkToSheet2 = $"'{dataSheet2.Name}'!A1";
            summarySheet.Hyperlinks.Add(2, 1, 1, 1, linkToSheet2); // B3 cell

            // -------------------------------------------------
            // 3. Save the workbook as PDF preserving hyperlinks
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Export document structure helps retain navigation links
            pdfOptions.ExportDocumentStructure = true;

            // Save to PDF file
            workbook.Save("WorkbookReport.pdf", pdfOptions);

            // Clean up
            workbook.Dispose();

            Console.WriteLine("PDF report generated with embedded worksheet hyperlinks.");
        }
    }
}