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

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive data");

            // Set the password that will be required to open the workbook
            workbook.Settings.Password = "UserDefinedPassword";

            // Apply AES‑256 encryption (key length 256 bits)
            // EncryptionType is ignored for modern Excel formats but must be provided
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the encrypted workbook
            workbook.Save("EncryptedWorkbook.xlsx");
        }
    }
}