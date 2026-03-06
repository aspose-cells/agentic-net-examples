using System;
using Aspose.Cells;

class EncryptWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sensitive Data");

        // Set a password – this triggers encryption when the workbook is saved
        workbook.Settings.Password = "StrongPassword123";

        // Define encryption options (type is ignored for .xlsx but required by the API)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Optionally protect the workbook structure with the same password
        workbook.Protect(ProtectionType.Structure, "StrongPassword123");

        // Save the encrypted workbook
        workbook.Save("EncryptedWorkbook.xlsx");
    }
}