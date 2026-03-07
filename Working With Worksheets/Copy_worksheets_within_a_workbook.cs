using System;
using Aspose.Cells;

namespace WorksheetCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Add a worksheet and put some sample data
            // -------------------------------------------------
            Worksheet original = workbook.Worksheets[0];
            original.Name = "Original";
            original.Cells["A1"].PutValue("This is the original sheet");
            original.Cells["B2"].PutValue(123);

            // -------------------------------------------------
            // 2. Copy the worksheet using AddCopy(string)
            // -------------------------------------------------
            int copiedIndexByName = workbook.Worksheets.AddCopy("Original");
            Worksheet copiedByName = workbook.Worksheets[copiedIndexByName];
            copiedByName.Name = "CopyByName";
            copiedByName.Cells["A1"].PutValue("This is the copy (by name)");

            // -------------------------------------------------
            // 3. Copy the worksheet using AddCopy(int)
            // -------------------------------------------------
            int copiedIndexByIndex = workbook.Worksheets.AddCopy(0); // index of "Original"
            Worksheet copiedByIndex = workbook.Worksheets[copiedIndexByIndex];
            copiedByIndex.Name = "CopyByIndex";
            copiedByIndex.Cells["A1"].PutValue("This is the copy (by index)");

            // -------------------------------------------------
            // 4. Copy multiple worksheets using AddCopy(Worksheet[], string[])
            // -------------------------------------------------
            // Add two more worksheets to serve as sources
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Data in Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
            sheet3.Cells["A1"].PutValue("Data in Sheet3");

            // Prepare source array and destination names
            Worksheet[] sourceSheets = new Worksheet[] { sheet2, sheet3 };
            string[] destNames = new string[] { "CopyOfSheet2", "CopyOfSheet3" };

            // Perform group copy
            workbook.Worksheets.AddCopy(sourceSheets, destNames);

            // -------------------------------------------------
            // 5. Copy a worksheet to another workbook using Worksheet.Copy with options
            // -------------------------------------------------
            // Create a second workbook
            Workbook destWorkbook = new Workbook();

            // Ensure the destination has at least one worksheet
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "Destination";

            // Configure copy options (e.g., keep formulas referring to same‑named sheets)
            CopyOptions copyOptions = new CopyOptions();
            copyOptions.ReferToSheetWithSameName = true;

            // Copy the original worksheet into the destination worksheet
            destSheet.Copy(original, copyOptions);
            destSheet.Name = "OriginalCopiedToOtherWorkbook";

            // -------------------------------------------------
            // 6. Save both workbooks
            // -------------------------------------------------
            workbook.Save("WorksheetCopyDemo.xlsx");
            destWorkbook.Save("WorksheetCopyDemo_SecondWorkbook.xlsx");

            Console.WriteLine("Worksheets copied successfully.");
        }
    }
}