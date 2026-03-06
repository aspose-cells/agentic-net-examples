using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive data");

        // Choose encryption type and key length (e.g., StrongCryptographicProvider with 128-bit key)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Set the password required to open the workbook
        workbook.Settings.Password = "MySecurePassword";

        // Save the encrypted workbook as XLSX
        workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}