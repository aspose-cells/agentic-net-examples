using System;
using Aspose.Cells;

class ProtectWorkbookStructure
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Protect only the workbook structure with a password
        // This prevents adding, deleting, renaming, or moving worksheets
        workbook.Protect(ProtectionType.Structure, "StrongPassword123");

        // Save the protected workbook to an XLSX file
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}