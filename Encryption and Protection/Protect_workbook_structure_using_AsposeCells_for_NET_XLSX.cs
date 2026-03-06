using System;
using Aspose.Cells;

class ProtectWorkbookStructure
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Protect only the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "myPassword123");

        // Save the protected workbook in XLSX format
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}