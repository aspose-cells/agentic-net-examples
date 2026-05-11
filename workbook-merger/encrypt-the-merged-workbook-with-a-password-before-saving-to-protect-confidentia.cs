using System;
using Aspose.Cells;

namespace WorkbookEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (merged workbook can be built here)
            Workbook workbook = new Workbook();

            // Example: add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Confidential Data");

            // Set the password that will be required to open the workbook
            workbook.Settings.Password = "StrongPassword123";

            // (Optional) Define encryption strength for older Excel formats
            // workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook to a file
            workbook.Save("MergedWorkbook_Encrypted.xlsx");

            // Verify that the workbook is encrypted
            Console.WriteLine("Workbook encrypted: " + workbook.Settings.IsEncrypted);
        }
    }
}