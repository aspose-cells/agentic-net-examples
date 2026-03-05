using System;
using System.Data;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExportRangeWithHeader
{
    class Program
    {
        static void Main()
        {
            // Create source workbook and fill data (including header)
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            sourceSheet.Cells["A1"].PutValue("Product");
            sourceSheet.Cells["B1"].PutValue("Category");
            sourceSheet.Cells["C1"].PutValue("Price");

            sourceSheet.Cells["A2"].PutValue("Laptop");
            sourceSheet.Cells["B2"].PutValue("Electronics");
            sourceSheet.Cells["C2"].PutValue(1200.50);

            sourceSheet.Cells["A3"].PutValue("Phone");
            sourceSheet.Cells["B3"].PutValue("Electronics");
            sourceSheet.Cells["C3"].PutValue(899.99);

            sourceSheet.Cells["A4"].PutValue("Desk");
            sourceSheet.Cells["B4"].PutValue("Furniture");
            sourceSheet.Cells["C4"].PutValue(250.00);

            // Define the range that includes the header (A1:C4)
            AsposeRange