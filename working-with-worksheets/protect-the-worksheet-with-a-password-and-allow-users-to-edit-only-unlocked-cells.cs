using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtectionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example data
            sheet.Cells["A1"].PutValue("Locked Cell");
            sheet.Cells["B1"].PutValue("Locked Cell");
            sheet.Cells["A2"].PutValue("Editable Cell");
            sheet.Cells["B2"].PutValue("Editable Cell");

            // Unlock the cells that should be editable when the sheet is protected
            // Cells A2 and B2 will be unlocked
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false; // Unlock the cells

            // Apply the unlocked style to the desired range
            sheet.Cells["A2"].SetStyle(unlockedStyle);
            sheet.Cells["B2"].SetStyle(unlockedStyle);

            // Configure protection options
            Protection protection = sheet.Protection;
            protection.AllowSelectingLockedCell = false;      // Do not allow selection of locked cells
            protection.AllowSelectingUnlockedCell = true;    // Allow selection of unlocked cells
            protection.Password = "SecurePass123";           // Set the password for the worksheet

            // Protect the worksheet with all protection types and the specified password
            sheet.Protect(ProtectionType.All, "SecurePass123", null);

            // Save the workbook
            workbook.Save("ProtectedWorksheet.xlsx");
        }
    }
}