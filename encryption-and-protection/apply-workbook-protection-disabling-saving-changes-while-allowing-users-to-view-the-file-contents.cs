using System;
using Aspose.Cells;

class WorkbookWriteProtectionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("This workbook is read‑only unless the password is provided.");

        // Configure write protection:
        // - Set a password required to modify the file
        // - Recommend opening the file as read‑only
        // - Optionally set the author name
        workbook.Settings.WriteProtection.Password = "modify123";
        workbook.Settings.WriteProtection.RecommendReadOnly = true;
        workbook.Settings.WriteProtection.Author = "Admin";

        // Save the protected workbook
        workbook.Save("ReadOnlyWorkbook.xlsx");

        // Clean up
        workbook.Dispose();
    }
}