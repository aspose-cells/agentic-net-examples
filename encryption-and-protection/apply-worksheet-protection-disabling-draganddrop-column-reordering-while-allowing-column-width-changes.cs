using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the worksheet protection settings
        Protection protection = sheet.Protection;

        // Allow column formatting (e.g., changing column width)
        protection.AllowFormattingColumn = true;

        // Disallow column insertion, deletion, and moving (drag‑and‑drop reordering)
        protection.AllowInsertingColumn = false;
        protection.AllowDeletingColumn = false;

        // Protect the worksheet with all protection options (no password)
        sheet.Protect(ProtectionType.All);

        // Save the protected workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}