using System;
using Aspose.Cells;

class ProtectWorksheet
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Configure protection settings to prevent accidental deletion
        Protection protection = sheet.Protection;
        protection.AllowDeletingRow = false;      // Disallow row deletion
        protection.AllowDeletingColumn = false;   // Disallow column deletion
        protection.AllowInsertingRow = false;     // Optional: also prevent row insertion
        protection.AllowInsertingColumn = false;  // Optional: also prevent column insertion

        // Apply protection to the worksheet (no password required)
        sheet.Protect(ProtectionType.All);

        // Save the protected workbook to an XLSX file
        workbook.Save("ProtectedList.xlsx");
    }
}