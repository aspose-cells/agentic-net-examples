using System;
using Aspose.Cells;

namespace AsposeCellsCopyMoveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Add initial worksheets and populate data
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
            sheet1.Name = "SourceSheet";
            sheet1.Cells["A1"].PutValue("Original Data");
            sheet1.Cells["A2"].PutValue(123);

            // Add a second sheet for later swapping
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("Second Sheet Data");

            // -------------------------------------------------
            // 2. Copy a worksheet using AddCopy (by name)
            // -------------------------------------------------
            int copiedIndex = workbook.Worksheets.AddCopy("SourceSheet"); // rule AddCopy(string)
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "CopiedByAddCopy";

            // -------------------------------------------------
            // 3. Copy a worksheet using Worksheet.Copy with options
            // -------------------------------------------------
            Worksheet destSheet = workbook.Worksheets.Add("CopyWithOptions");
            CopyOptions copyOptions = new CopyOptions();
            copyOptions.ReferToSheetWithSameName = true; // keep formula references within same workbook
            destSheet.Copy(sheet1, copyOptions); // rule Worksheet.Copy(Worksheet, CopyOptions)

            // -------------------------------------------------
            // 4. Move a worksheet to a new position
            // -------------------------------------------------
            // Move "CopiedByAddCopy" to index 0 (make it the first sheet)
            copiedSheet.MoveTo(0); // rule Worksheet.MoveTo(int)

            // -------------------------------------------------
            // 5. Swap two worksheets
            // -------------------------------------------------
            // Swap the sheet now at index 1 with the sheet at index 2
            workbook.Worksheets.SwapSheet(1, 2); // rule WorksheetCollection.SwapSheet(int, int)

            // -------------------------------------------------
            // 6. Insert a new worksheet at a specific index
            // -------------------------------------------------
            Worksheet inserted = workbook.Worksheets.Insert(2, SheetType.Worksheet, "InsertedSheet"); // rule Insert(int, SheetType, string)
            inserted.Cells["A1"].PutValue("Inserted sheet content");

            // -------------------------------------------------
            // 7. Remove a worksheet by name
            // -------------------------------------------------
            workbook.Worksheets.RemoveAt("SecondSheet"); // rule RemoveAt(string)

            // -------------------------------------------------
            // 8. Save the workbook (lifecycle: save)
            // -------------------------------------------------
            workbook.Save("CopyMoveDemo.xlsx");
        }
    }
}