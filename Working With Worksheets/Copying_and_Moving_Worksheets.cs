using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class WorksheetCopyMoveDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the default first worksheet and set up initial data
            Worksheet original = workbook.Worksheets[0];
            original.Name = "Original";
            original.Cells["A1"].PutValue("Original Data");

            // 1. Copy the worksheet using AddCopy(string)
            int copiedIndex = workbook.Worksheets.AddCopy("Original"); // returns index of new sheet
            Worksheet copy1 = workbook.Worksheets[copiedIndex];
            copy1.Name = "Copy1";
            copy1.Cells["A2"].PutValue("Data in copied sheet");

            // 2. Add a new worksheet normally
            Worksheet second = workbook.Worksheets.Add("Second");
            second.Cells["A1"].PutValue("Second Sheet Data");

            // 3. Copy contents of a worksheet using Worksheet.Copy with CopyOptions
            CopyOptions copyOptions = new CopyOptions();
            copyOptions.ReferToSheetWithSameName = true; // keep formula references within same workbook
            Worksheet copyWithOptions = workbook.Worksheets.Add("CopyWithOptions");
            copyWithOptions.Copy(original, copyOptions);
            copyWithOptions.Cells["A3"].PutValue("Copied with options");

            // 4. Insert a worksheet at a specific position
            Worksheet inserted = workbook.Worksheets.Insert(1, SheetType.Worksheet, "Inserted");
            inserted.Cells["A1"].PutValue("Inserted Sheet");

            // 5. Move a worksheet to a new index using MoveTo
            inserted.MoveTo(4);

            // 6. Swap two worksheets using SwapSheet
            workbook.Worksheets.SwapSheet(0, 4);

            // 7. Remove a worksheet by name
            int secondIndex = workbook.Worksheets["Second"].Index;
            workbook.Worksheets.RemoveAt(secondIndex);

            // Save the workbook
            workbook.Save("WorksheetCopyMoveDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main()
        {
            WorksheetCopyMoveDemo.Run();
        }
    }
}