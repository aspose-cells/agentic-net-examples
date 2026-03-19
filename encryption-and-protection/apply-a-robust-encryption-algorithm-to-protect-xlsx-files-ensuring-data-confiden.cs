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

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Set a strong password for opening the workbook
            workbook.Settings.Password = "StrongPassword123!";

            // Apply strong encryption (AES 256) – Excel 2007/2010+ uses SHA AES internally.
            // The EncryptionType is ignored for newer formats, but we set it for completeness.
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Protect the workbook structure to prevent adding/removing sheets without the password
            workbook.Protect(ProtectionType.Structure, "StrongPassword123!");

            // Save the encrypted workbook as XLSX
            workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // Verify encryption by attempting to load the file with the password
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "StrongPassword123!"
            };
            Workbook loadedWorkbook = new Workbook("EncryptedWorkbook.xlsx", loadOptions);

            // Output a cell value to confirm successful decryption
            Console.WriteLine("Decrypted cell A1 value: " + loadedWorkbook.Worksheets[0].Cells["A1"].StringValue);
        }
    }
}