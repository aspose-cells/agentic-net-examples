using System;
using System.IO;
using Aspose.Cells;

class EncryptionDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data
        workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive data");

        // Select encryption algorithm and key length
        workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

        // Set the password for opening the workbook
        string password = "MySecretPassword";
        workbook.Settings.Password = password;

        // Save the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";
        workbook.Save(filePath);

        // Validate the password against the saved file
        bool isPasswordValid;
        using (FileStream stream = File.OpenRead(filePath))
        {
            isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
        }

        Console.WriteLine($"Password validation result: {isPasswordValid}");
    }
}