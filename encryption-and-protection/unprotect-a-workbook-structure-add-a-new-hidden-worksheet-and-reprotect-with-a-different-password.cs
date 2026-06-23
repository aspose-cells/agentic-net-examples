using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Initially protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "oldPassword");

        // Unprotect the workbook structure using the original password
        workbook.Unprotect("oldPassword");

        // Add a new worksheet and hide it
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet hiddenSheet = workbook.Worksheets[newSheetIndex];
        hiddenSheet.IsVisible = false; // make the worksheet hidden

        // Re‑protect the workbook structure with a different password
        workbook.Protect(ProtectionType.Structure, "newPassword");

        // Save the modified workbook
        workbook.Save("ModifiedWorkbook.xlsx");
    }
}