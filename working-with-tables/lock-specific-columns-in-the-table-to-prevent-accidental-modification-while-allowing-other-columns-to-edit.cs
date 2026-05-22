using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class LockColumnsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a StyleFlag indicating the Locked attribute should be applied
            StyleFlag flag = new StyleFlag { Locked = true };

            // Unlock all columns first (so only selected columns will be locked later)
            for (int col = 0; col <= worksheet.Cells.MaxColumn; col++)
            {
                Style style = worksheet.Cells.Columns[col].GetStyle();
                style.IsLocked = false;
                worksheet.Cells.Columns[col].ApplyStyle(style, flag);
            }

            // Columns to lock (e.g., column A and C) – 0‑based indexes
            int[] columnsToLock = { 0, 2 };

            foreach (int colIndex in columnsToLock)
            {
                Style style = worksheet.Cells.Columns[colIndex].GetStyle();
                style.IsLocked = true;
                worksheet.Cells.Columns[colIndex].ApplyStyle(style, flag);
            }

            // Protect the worksheet so that the locking takes effect
            worksheet.Protect(ProtectionType.All);
            // Allow selecting unlocked cells while preventing edits on locked ones
            worksheet.Protection.AllowSelectingUnlockedCell = true;

            // Save the workbook
            string outputPath = "LockedColumnsDemo.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}