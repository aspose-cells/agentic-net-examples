using System;
using Aspose.Cells;

class ProtectWorkbookStructure
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook's structure with a custom password
        workbook.Protect(ProtectionType.Structure, "MySecretPassword");

        // Save the protected workbook to a file
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}