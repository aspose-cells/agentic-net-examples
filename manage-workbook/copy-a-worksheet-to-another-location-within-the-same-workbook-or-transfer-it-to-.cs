using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetCopyDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a source workbook and add sample data
            // -------------------------------------------------
            Workbook sourceWorkbook = new Workbook();                     // create a new workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];        // get the first worksheet
            sourceSheet.Name = "SourceSheet";
            sourceSheet.Cells["A1"].PutValue("Hello");
            sourceSheet.Cells["B1"].PutValue(123);
            sourceSheet.Cells["A2"].PutValue("World");
            sourceSheet.Cells["B2"].PutValue(456);

            // -------------------------------------------------
            // 2. Copy the worksheet to a new location within the same workbook
            //    using WorksheetCollection.AddCopy (by index)
            // -------------------------------------------------
            int copiedIndex = sourceWorkbook.Worksheets.AddCopy(0);      // copy the first sheet
            Worksheet copiedSheet = sourceWorkbook.Worksheets[copiedIndex];
            copiedSheet.Name = "CopiedWithinSameWorkbook";

            // -------------------------------------------------
            // 3. Move the copied sheet to position 0 (make it the first sheet)
            // -------------------------------------------------
            copiedSheet.MoveTo(0);                                       // now the copied sheet is first

            // -------------------------------------------------
            // 4. Save the workbook that contains the original and copied sheets
            // -------------------------------------------------
            sourceWorkbook.Save("WorkbookWithCopiedSheet.xlsx");

            // -------------------------------------------------
            // 5. Create a destination workbook (empty) and copy the original sheet into it
            //    using Worksheet.Copy(sourceSheet, copyOptions)
            // -------------------------------------------------
            Workbook destWorkbook = new Workbook();                       // create an empty workbook
            Worksheet destSheet = destWorkbook.Worksheets[0];            // get its first (default) sheet

            // Configure copy options if needed (e.g., keep formulas referring to destination sheet)
            CopyOptions options = new CopyOptions
            {
                ReferToDestinationSheet = true   // optional, demonstrates usage of CopyOptions
            };

            // Copy the source sheet into the destination sheet
            destSheet.Copy(sourceSheet, options);
            destSheet.Name = "CopiedFromSourceWorkbook";

            // -------------------------------------------------
            // 6. Save the destination workbook
            // -------------------------------------------------
            destWorkbook.Save("CopiedToAnotherWorkbook.xlsx");

            Console.WriteLine("Worksheet copy operations completed successfully.");
        }
    }
}