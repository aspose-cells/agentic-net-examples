using System;
using Aspose.Cells;

class ProtectUnprotectWorkbook
{
    static void Main()
    {
        // Paths for the original, protected, and unprotected files
        string inputPath = "input.xlsx";
        string protectedPath = "protected.xlsx";
        string unprotectedPath = "unprotected.xlsx";

        // Password to protect/unprotect the workbook
        string password = "myPassword";

        // Load the existing workbook
        Workbook wb = new Workbook(inputPath);

        // Protect the workbook (structure and windows) with a password
        wb.Protect(ProtectionType.All, password);

        // Save the protected workbook
        wb.Save(protectedPath, SaveFormat.Xlsx);

        // Load the protected workbook
        Workbook protectedWb = new Workbook(protectedPath);

        // Unprotect the workbook using the same password
        protectedWb.Unprotect(password);

        // Save the unprotected workbook
        protectedWb.Save(unprotectedPath, SaveFormat.Xlsx);
    }
}