using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted content");

            // Set a password to protect the workbook (WorkbookSettings.Password property)
            workbook.Settings.Password = "StrongPassword123";

            // Define encryption options (Workbook.SetEncryptionOptions method)
            // Using StrongCryptographicProvider with a 128‑bit key (EncryptionType enum)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook as XLSX (lifecycle rule: save)
            workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);

            // Optional: Verify that the file is encrypted
            FileFormatInfo info = FileFormatUtil.DetectFileFormat("EncryptedWorkbook.xlsx");
            Console.WriteLine("Is the saved file encrypted? " + info.IsEncrypted);
        }
    }
}