using System;
using Aspose.Cells;

namespace AsposeCellsSheetOperations
{
    class Program
    {
        static void Main(string[] args)
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add sample worksheets
            // -------------------------------------------------
            Workbook workbook = new Workbook(); // create new workbook

            // Default first sheet is already present; rename it
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Original";

            // Add two more sheets
            Worksheet sheet2 = workbook.Worksheets.Add("Data");
            Worksheet sheet3 = workbook.Worksheets.Add("Summary");

            // Populate some data
            sheet1.Cells["A1"].PutValue("This is the original sheet");
            sheet2.Cells["A1"].PutValue("Data sheet header");
            sheet2.Cells["A2"].PutValue(123);
            sheet3.Cells["A1"].PutValue("Summary sheet");
            sheet3.Cells["B2"].Formula = "=SUM(Data!A2)"; // reference to sheet2

            // -------------------------------------------------
            // 2. Copy a sheet using WorksheetCollection.AddCopy(string)
            // -------------------------------------------------
            // This creates a new sheet that is a copy of "Original"
            int copiedIndex = workbook.Worksheets.AddCopy("Original");
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "Original_Copy";

            // Modify copied sheet to prove it's independent
            copiedSheet.Cells["A1"].PutValue("This is the copied sheet");

            // -------------------------------------------------
            // 3. Copy a sheet using Worksheet.Copy with CopyOptions
            // -------------------------------------------------
            // We'll copy "Data" sheet to a new sheet and keep formula references
            Worksheet dataCopy = workbook.Worksheets.Add("Data_Copy");
            CopyOptions copyOpts = new CopyOptions
            {
                // Keep formulas referring to the original sheet name (default true)
                ReferToSheetWithSameName = true
            };
            dataCopy.Copy(sheet2, copyOpts);

            // -------------------------------------------------
            // 4. Move a sheet to a different position using Worksheet.MoveTo
            // -------------------------------------------------
            // Move "Summary" sheet to index 1 (second position)
            sheet3.MoveTo(1); // after this, order: Original, Summary, Original_Copy, Data, Data_Copy

            // -------------------------------------------------
            // 5. Swap two sheets using WorksheetCollection.SwapSheet
            // -------------------------------------------------
            // Swap "Original" (index 0) with "Data_Copy" (current index 4)
            workbook.Worksheets.SwapSheet(0, 4); // order changes accordingly

            // -------------------------------------------------
            // 6. Save the workbook to demonstrate the result
            // -------------------------------------------------
            workbook.Save("SheetOperationsResult.xlsx");

            Console.WriteLine("Workbook created and saved as SheetOperationsResult.xlsx");
        }
    }
}