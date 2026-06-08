using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Protected and Encrypted");

        // Protect the workbook structure with a password
        workbook.Protect(ProtectionType.Structure, "protectPwd");

        // Set a password required to open the workbook (encryption)
        workbook.Settings.Password = "openPwd";

        // Use strong encryption (optional but recommended)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the workbook in XLSX format
        workbook.Save("ProtectedEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}