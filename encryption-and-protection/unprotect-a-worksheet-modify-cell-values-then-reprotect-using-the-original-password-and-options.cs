using Aspose.Cells;
using System;

class WorksheetUnprotectModifyProtect
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some initial data
        sheet.Cells["A1"].PutValue("Original");
        sheet.Cells["B1"].PutValue(123);

        // ----- Set up protection options -----
        Protection protection = sheet.Protection;
        protection.AllowEditingContent = false;          // Disallow editing locked cells
        protection.AllowEditingObject = false;           // Disallow editing objects
        protection.AllowFormattingCell = true;           // Allow cell formatting
        protection.Password = "mySecret";                // Set password

        // Apply protection (all protection types) with the password
        sheet.Protect(ProtectionType.All, protection.Password, null);

        // ----- Preserve original password and options -----
        string originalPassword = protection.Password;
        bool allowEditingContent = protection.AllowEditingContent;
        bool allowEditingObject = protection.AllowEditingObject;
        bool allowFormattingCell = protection.AllowFormattingCell;
        // Add more option variables here if you need to preserve additional settings

        // ----- Unprotect the worksheet using the original password -----
        sheet.Unprotect(originalPassword);

        // ----- Modify cell values while the sheet is unprotected -----
        sheet.Cells["A1"].PutValue("Modified");
        sheet.Cells["B1"].PutValue(456);

        // ----- Restore protection options -----
        protection = sheet.Protection;
        protection.AllowEditingContent = allowEditingContent;
        protection.AllowEditingObject = allowEditingObject;
        protection.AllowFormattingCell = allowFormattingCell;
        protection.Password = originalPassword;

        // Re‑protect the worksheet with the same password and options
        sheet.Protect(ProtectionType.All, originalPassword, null);

        // Save the workbook
        workbook.Save("Result.xlsx");
    }
}