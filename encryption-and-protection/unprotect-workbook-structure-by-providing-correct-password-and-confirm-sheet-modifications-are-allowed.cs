using System;
using Aspose.Cells;

class UnprotectWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook and protect its structure with a password
        Workbook workbook = new Workbook();
        workbook.Protect(ProtectionType.Structure, "securePwd");
        Console.WriteLine("Workbook initially protected: " + workbook.IsWorkbookProtectedWithPassword);
        workbook.Save("protected_workbook.xlsx");

        // Load the protected workbook
        Workbook loadedWorkbook = new Workbook("protected_workbook.xlsx");
        Console.WriteLine("Loaded workbook protected: " + loadedWorkbook.IsWorkbookProtectedWithPassword);

        // Unprotect the workbook structure using the correct password
        loadedWorkbook.Unprotect("securePwd");
        Console.WriteLine("Workbook after unprotect: " + loadedWorkbook.IsWorkbookProtectedWithPassword);

        // Confirm that modifications are now allowed: add a new worksheet and write a value
        int newSheetIndex = loadedWorkbook.Worksheets.Add();
        Worksheet newSheet = loadedWorkbook.Worksheets[newSheetIndex];
        newSheet.Name = "ModifiedSheet";
        newSheet.Cells["A1"].PutValue("Modifications are allowed after unprotect.");
        Console.WriteLine("Added new worksheet and wrote to cell A1.");

        // Save the unprotected workbook
        loadedWorkbook.Save("unprotected_workbook.xlsx");
    }
}