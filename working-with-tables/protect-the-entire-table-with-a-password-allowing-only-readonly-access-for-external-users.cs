using Aspose.Cells;

class ProtectTableDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Populate some data in the worksheet
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue("Data1");
        worksheet.Cells["B2"].PutValue("Data2");

        // Configure worksheet protection to allow only read‑only access
        Protection protection = worksheet.Protection;
        protection.AllowEditingContent = false;          // Disallow editing cell contents
        protection.AllowEditingObject = false;           // Disallow editing objects
        protection.AllowEditingScenario = false;         // Disallow editing scenarios
        protection.AllowFormattingCell = false;          // Disallow formatting cells
        protection.AllowFormattingColumn = false;        // Disallow formatting columns
        protection.AllowFormattingRow = false;           // Disallow formatting rows
        protection.AllowInsertingColumn = false;         // Disallow inserting columns
        protection.AllowInsertingRow = false;            // Disallow inserting rows
        protection.AllowInsertingHyperlink = false;      // Disallow inserting hyperlinks
        protection.AllowDeletingColumn = false;          // Disallow deleting columns
        protection.AllowDeletingRow = false;             // Disallow deleting rows
        protection.AllowSorting = false;                 // Disallow sorting
        protection.AllowFiltering = false;               // Disallow filtering
        protection.AllowUsingPivotTable = false;         // Disallow using pivot tables
        protection.AllowSelectingLockedCell = true;      // Allow selecting locked cells (read‑only)
        protection.AllowSelectingUnlockedCell = true;    // Allow selecting unlocked cells

        // Set a password so that only users with the password can unprotect the sheet
        protection.Password = "ReadOnlyPwd";

        // Save the workbook
        workbook.Save("ReadOnlyProtectedTable.xlsx");
    }
}