using System;
using Aspose.Cells;

class ProtectMergedWorkbook
{
    static void Main()
    {
        // Load the workbook that resulted from the combine operation
        Workbook mergedWorkbook = new Workbook("merged.xlsx");

        // Protect the workbook's structure (prevents adding, deleting, renaming sheets) with a password
        mergedWorkbook.Protect(ProtectionType.Structure, "MySecretPassword");

        // Save the password‑protected workbook
        mergedWorkbook.Save("merged_protected.xlsx", SaveFormat.Xlsx);
    }
}