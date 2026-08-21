// Title: Export Multiple Excel Tables from Separate Worksheets to a Single PDF with Aspose.Cells for .NET (C#)
// Description: Creates a workbook with two ListObject tables on different worksheets, configures PdfSaveOptions (OnePagePerSheet & AllColumnsInOnePagePerSheet) and saves the whole workbook as one PDF while keeping each table’s layout intact.
// Keywords: Aspose.Cells export multiple tables PDF | C# PdfSaveOptions OnePagePerSheet | Aspose.Cells ListObject to PDF | export Excel worksheets to single PDF .NET | fit columns one page per sheet Aspose | multiple worksheets PDF conversion | Aspose.Cells PDF export example | C# generate PDF from Excel tables
// Common Searches: how to export several worksheets with tables to one PDF using Aspose.Cells | Aspose.Cells OnePagePerSheet example C# | save Excel ListObject as single PDF file | export multiple Excel tables to PDF .NET | Aspose.Cells PDFSaveOptions all columns in one page
// Developer Intent: Generate a single PDF that contains each worksheet’s Excel table on its own page, preserving column widths and table formatting.
// Use Cases: Combine customer and order tables from separate sheets into one printable PDF report. | Create page‑wise PDF invoices where each sheet represents a different section of the document. | Automate PDF generation in an ASP.NET service that sends a consolidated data file via email.
// AI Prompts: Write C# code with Aspose.Cells to export three worksheets, each containing a ListObject, to a single PDF with one page per sheet. | Explain how PdfSaveOptions.OnePagePerSheet and AllColumnsInOnePagePerSheet affect the layout of exported tables. | Provide a step‑by‑step tutorial for adding ListObjects to ranges and saving the workbook as a PDF using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExportMultipleTables
{
    // Creates a workbook with two ListObject tables on different worksheets, configures PdfSaveOptions (OnePagePerSheet & AllColumnsInOnePagePerSheet) and saves the whole workbook as one PDF while keeping each table’s layout intact.
    class Program
    {
        static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // First worksheet – a table of customers.
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Customers";

            // Header row.
            sheet1.Cells["A1"].PutValue("ID");
            sheet1.Cells["B1"].PutValue("Name");
            sheet1.Cells["C1"].PutValue("Country");

            // Sample data rows.
            for (int i = 2; i <= 6; i++)
            {
                sheet1.Cells[$"A{i}"].PutValue(i - 1);                         // ID
                sheet1.Cells[$"B{i}"].PutValue($"Customer {i - 1}");          // Name
                sheet1.Cells[$"C{i}"].PutValue(i % 2 == 0 ? "USA" : "UK");    // Country
            }

            // Convert the range into an Excel table (ListObject).
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            sheet1.ListObjects.Add(0, 0, 6, 3, true);

            // -------------------------------------------------
            // Second worksheet – a table of orders.
            // -------------------------------------------------
            Worksheet sheet2 = workbook.Worksheets.Add("Orders");

            // Header row.
            sheet2.Cells["A1"].PutValue("OrderID");
            sheet2.Cells["B1"].PutValue("CustomerID");
            sheet2.Cells["C1"].PutValue("Amount");

            // Sample data rows.
            for (int i = 2; i <= 8; i++)
            {
                sheet2.Cells[$"A{i}"].PutValue(1000 + i);          // OrderID
                sheet2.Cells[$"B{i}"].PutValue(i - 1);            // CustomerID
                sheet2.Cells[$"C{i}"].PutValue((i - 1) * 123.45); // Amount
            }

            // Convert the range into an Excel table (ListObject).
            sheet2.ListObjects.Add(0, 0, 8, 3, true);

            // -------------------------------------------------
            // PDF save options – keep each worksheet on a single page
            // and fit all columns horizontally.
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,
                AllColumnsInOnePagePerSheet = true
            };

            // Export the entire workbook (both tables) to a single PDF file.
            workbook.Save("MultipleTables.pdf", pdfOptions);
        }
    }
}
