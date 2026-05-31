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

            // Example data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data");

            // Protect the workbook structure with a password
            string workbookPassword = "wbStructurePwd";
            workbook.Protect(ProtectionType.Structure, workbookPassword);

            // Protect all worksheets (no password, full protection)
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.Protect(ProtectionType.All);
            }

            // Set a password to encrypt the file when saved
            workbook.Settings.Password = "fileOpenPwd";

            // Save the workbook as an encrypted XLSX file
            workbook.Save("ProtectedEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}