using System;
using Aspose.Cells;

class RC4EncryptionDemo
{
    static void Main()
    {
        // Create a new workbook and put some data in the first cell
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive Information");

        // Set a password for the workbook (this will trigger encryption)
        workbook.Settings.Password = "rc4password";

        // Specify encryption options (using StrongCryptographicProvider as the closest option)
        // RC4 is not directly exposed; this demonstrates encryption with a password.
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook (Excel 97-2003 format)
        string encryptedFile = "EncryptedWorkbook.xls";
        workbook.Save(encryptedFile, SaveFormat.Excel97To2003);

        // Load the encrypted workbook using the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = "rc4password";
        Workbook loadedWorkbook = new Workbook(encryptedFile, loadOptions);

        // Verify that the data is consistent after decryption
        string decryptedValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Decrypted cell value: " + decryptedValue);
    }
}