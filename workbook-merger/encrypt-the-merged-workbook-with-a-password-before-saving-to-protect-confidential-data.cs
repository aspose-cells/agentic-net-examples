using System;
using Aspose.Cells;

class EncryptWorkbookExample
{
    static void Main()
    {
        // Load the merged workbook (replace with the actual file path)
        string mergedFilePath = "mergedWorkbook.xlsx";
        Workbook workbook = new Workbook(mergedFilePath);

        // Set a password to encrypt the workbook
        workbook.Settings.Password = "StrongPassword123";

        // Optional: define encryption algorithm and key length (AES 128)
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook
        workbook.Save("mergedWorkbook_encrypted.xlsx");
    }
}