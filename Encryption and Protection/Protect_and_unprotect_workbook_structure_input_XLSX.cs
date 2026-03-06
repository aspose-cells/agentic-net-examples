using System;
using Aspose.Cells;

class ProtectUnprotectWorkbook
{
    static void Main()
    {
        // Load the existing workbook (input XLSX)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "myPassword");

        // Save the protected workbook
        string protectedPath = "protected.xlsx";
        workbook.Save(protectedPath, SaveFormat.Xlsx);

        // Load the protected workbook to verify protection status
        Workbook protectedWorkbook = new Workbook(protectedPath);
        bool isProtected = protectedWorkbook.Settings.IsProtected;
        Console.WriteLine("Workbook is protected: " + isProtected);

        // Unprotect the workbook using the same password
        protectedWorkbook.Unprotect("myPassword");

        // Save the unprotected workbook
        string unprotectedPath = "unprotected.xlsx";
        protectedWorkbook.Save(unprotectedPath, SaveFormat.Xlsx);
    }
}