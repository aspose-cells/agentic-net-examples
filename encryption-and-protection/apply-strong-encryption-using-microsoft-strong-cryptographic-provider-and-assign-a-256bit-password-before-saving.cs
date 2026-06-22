using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Information");

            // Assign a password that will be used for encryption
            workbook.Settings.Password = "MyStrong256BitPassword";

            // Apply strong encryption (Microsoft Strong Cryptographic Provider) with a 256‑bit key
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the encrypted workbook
            workbook.Save("StrongEncryptedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}