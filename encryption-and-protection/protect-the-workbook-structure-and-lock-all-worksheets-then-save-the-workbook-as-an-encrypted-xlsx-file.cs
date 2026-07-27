using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "WorkbookStructurePwd");

            // Protect (lock) every worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Protect the worksheet without a password (all protection types)
                sheet.Protect(ProtectionType.All);
            }

            // Set a password to encrypt the workbook file (required to open)
            workbook.Settings.Password = "OpenFilePwd";

            // Save the workbook as an encrypted XLSX file
            workbook.Save("ProtectedEncryptedWorkbook.xlsx", SaveFormat.Xlsx);

            // Clean up
            workbook.Dispose();
        }
    }
}