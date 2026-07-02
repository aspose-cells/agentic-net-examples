using System;
using Aspose.Cells;

namespace AsposeCellsColumnLockDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Unlock all columns (set IsLocked = false) so they can be edited after protection
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;
            StyleFlag unlockedFlag = new StyleFlag();
            unlockedFlag.Locked = true; // Apply the Locked flag

            // Apply the unlocked style to all columns (0 to 255)
            for (int col = 0; col <= 255; col++)
            {
                sheet.Cells.Columns[col].ApplyStyle(unlockedStyle, unlockedFlag);
            }

            // Define which columns should remain locked (e.g., column 0 and column 2)
            int[] lockedColumns = new int[] { 0, 2 };
            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true;
            StyleFlag lockedFlag = new StyleFlag();
            lockedFlag.Locked = true; // Apply the Locked flag

            // Apply the locked style to the specified columns
            foreach (int colIndex in lockedColumns)
            {
                sheet.Cells.Columns[colIndex].ApplyStyle(lockedStyle, lockedFlag);
            }

            // Protect the worksheet so that the locking takes effect
            sheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("LockedColumnsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}