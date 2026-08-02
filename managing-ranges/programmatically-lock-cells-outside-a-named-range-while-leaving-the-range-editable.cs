using System;
using Aspose.Cells;

namespace AsposeCellsLockOutsideRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example data (optional)
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["C1"].PutValue("Header3");
            cells["A2"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["C2"].PutValue(30);
            cells["A3"].PutValue(40);
            cells["B3"].PutValue(50);
            cells["C3"].PutValue(60);

            // Ensure all cells are locked (default is true, but we set it explicitly for clarity)
            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true;
            StyleFlag lockedFlag = new StyleFlag();
            lockedFlag.Locked = true;
            cells.ApplyStyle(lockedStyle, lockedFlag);

            // Define the editable range (e.g., A2:C3)
            int startRow = 1;      // zero‑based index (row 2)
            int startColumn = 0;   // column A
            int endRow = 2;        // row 3
            int endColumn = 2;     // column C

            // Add the range to AllowEditRanges collection
            int rangeIndex = sheet.AllowEditRanges.Add("EditableRange", startRow, startColumn, endRow, endColumn);
            ProtectedRange editableRange = sheet.AllowEditRanges[rangeIndex];
            // Optional: set a password for the editable range
            editableRange.Password = "rangePwd";

            // Protect the worksheet (all protection types) with a sheet password
            sheet.Protect(ProtectionType.All, "sheetPwd", null);

            // Save the workbook
            workbook.Save("LockedOutsideRange.xlsx");
        }
    }
}