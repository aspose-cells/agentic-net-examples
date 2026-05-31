using System;
using Aspose.Cells;

class ExportMultipleTablesToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "Data";

        // -------------------- First Table --------------------
        // Populate data for the first table (rows 0-9, columns A-B)
        for (int i = 0; i < 10; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Item {i + 1}");
            worksheet.Cells[i, 1].PutValue(i * 10);
        }

        // Add the first table (ListObject) covering the populated range
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
        int firstTableIdx = worksheet.ListObjects.Add(0, 0, 9, 1, true);
        worksheet.ListObjects[firstTableIdx].DisplayName = "FirstTable";

        // -------------------- Second Table --------------------
        // Populate data for the second table (rows 12-19, columns D-E)
        for (int i = 0; i < 8; i++)
        {
            worksheet.Cells[12 + i, 3].PutValue($"Product {i + 1}");
            worksheet.Cells[12 + i, 4].PutValue(i * 5);
        }

        // Add the second table (ListObject) covering its range
        int secondTableIdx = worksheet.ListObjects.Add(12, 3, 19, 4, true);
        worksheet.ListObjects[secondTableIdx].DisplayName = "SecondTable";

        // -------------------- PDF Save Options --------------------
        // Configure options to preserve the layout of each table.
        // OnePagePerSheet = false allows the sheet to span multiple pages if needed.
        // AllColumnsInOnePagePerSheet = false keeps column widths as defined.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = false,
            AllColumnsInOnePagePerSheet = false
        };

        // Save the workbook as a single PDF file containing both tables.
        workbook.Save("MultipleTables.pdf", pdfOptions);
    }
}