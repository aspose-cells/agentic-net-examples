using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetTransferDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (source) and add sample data
            // -------------------------------------------------
            Workbook sourceWorkbook = new Workbook();                     // create
            Worksheet srcSheet = sourceWorkbook.Worksheets[0];
            srcSheet.Name = "DataSheet";

            // Populate some data and a formula that references another sheet
            srcSheet.Cells["A1"].PutValue("Header");
            srcSheet.Cells["A2"].PutValue(10);
            srcSheet.Cells["A3"].PutValue(20);
            srcSheet.Cells["B2"].Formula = "=SUM(A2:A3)";                // simple formula

            // Add a second sheet that will be referenced by the formula above
            Worksheet refSheet = sourceWorkbook.Worksheets.Add("RefSheet");
            refSheet.Cells["A1"].PutValue(5);

            // -------------------------------------------------
            // 2. Copy a worksheet within the same workbook using AddCopy
            // -------------------------------------------------
            // This creates a duplicate of "DataSheet" preserving formulas and formatting
            int copiedWithinIndex = sourceWorkbook.Worksheets.AddCopy("DataSheet");
            Worksheet copiedWithin = sourceWorkbook.Worksheets[copiedWithinIndex];
            copiedWithin.Name = "DataSheet_Copy";

            // -------------------------------------------------
            // 3. Move the newly copied sheet to position 0 (first tab)
            // -------------------------------------------------
            copiedWithin.MoveTo(0);   // MoveTo moves the sheet to the specified index

            // -------------------------------------------------
            // 4. Copy a worksheet from source workbook to a new destination workbook
            // -------------------------------------------------
            Workbook destWorkbook = new Workbook();          // create destination workbook
            destWorkbook.Worksheets.Clear();                // remove the default empty sheet

            // Prepare copy options to keep formula references pointing to the source sheet name
            CopyOptions copyOptions = new CopyOptions
            {
                ReferToSheetWithSameName = true   // ensures formulas still refer to "DataSheet"
            };

            // Add a blank sheet to destination with the same name as the source sheet
            Worksheet destSheet = destWorkbook.Worksheets.Add("DataSheet");

            // Copy contents, formats and formulas from source sheet to destination sheet
            destSheet.Copy(srcSheet, copyOptions);

            // -------------------------------------------------
            // 5. Save both workbooks to verify results
            // -------------------------------------------------
            sourceWorkbook.Save("SourceWorkbook_WithCopyAndMove.xlsx");
            destWorkbook.Save("DestinationWorkbook_CopiedSheet.xlsx");

            Console.WriteLine("Operations completed successfully.");
        }
    }
}