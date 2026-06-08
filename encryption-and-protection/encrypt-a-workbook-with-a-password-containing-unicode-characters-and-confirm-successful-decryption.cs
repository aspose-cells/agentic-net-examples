using System;
using Aspose.Cells;

class UnicodePasswordEncryptionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Unicode password test");

        // Define a password containing Unicode characters (Chinese characters and an emoji)
        string unicodePassword = "密码🔒";

        // Apply the password to encrypt the workbook
        wb.Settings.Password = unicodePassword;

        // Optionally set strong encryption options (128‑bit key)
        wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Save the encrypted workbook to a file
        string filePath = "UnicodeEncrypted.xlsx";
        wb.Save(filePath);

        // Verify that the workbook reports being encrypted
        Console.WriteLine("IsEncrypted after save: " + wb.Settings.IsEncrypted);

        // Load the encrypted workbook using the same Unicode password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = unicodePassword;
        Workbook loadedWb = new Workbook(filePath, loadOptions);

        // Confirm successful decryption by reading the previously saved cell value
        string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Decrypted cell value: " + cellValue);
    }
}