using System;
using Aspose.Cells;

class ProtectWorksheetDemo
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet you want to protect (first worksheet in this example)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the protection object for the worksheet
        Protection protection = worksheet.Protection;

        // Restrict editing of locked cells
        protection.AllowEditingContent = false;

        // Restrict formatting operations
        protection.AllowFormattingCell = false;
        protection.AllowFormattingColumn = false;
        protection.AllowFormattingRow = false;

        // Additional optional restrictions
        protection.AllowInsertingRow = false;
        protection.AllowInsertingColumn = false;
        protection.AllowDeletingRow = false;
        protection.AllowDeletingColumn = false;
        protection.AllowSorting = false;
        protection.AllowFiltering = false;

        // Set a password for the protection (optional)
        protection.Password = "pwd123";

        // Apply protection to the worksheet using all protection types
        worksheet.Protect(ProtectionType.All, protection.Password, null);

        // Save the workbook with the protected worksheet
        workbook.Save("output.xlsx");
    }
}