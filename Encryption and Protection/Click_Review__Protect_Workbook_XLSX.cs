using System;
using Aspose.Cells;

class ProtectWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook's structure with a password
        workbook.Protect(ProtectionType.Structure, "myPassword123");

        // Save the protected workbook in XLSX format
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}