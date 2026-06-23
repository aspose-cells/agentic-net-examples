using System;
using System.IO;
using Aspose.Cells;

class VerifyEncryptionPasswordDemo
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        // Password to verify
        string passwordToTest = "testPassword";

        // Verify the password using FileFormatUtil.VerifyPassword
        bool isPasswordCorrect;
        using (Stream stream = File.OpenRead(filePath))
        {
            isPasswordCorrect = FileFormatUtil.VerifyPassword(stream, passwordToTest);
        }

        // Log the verification result
        Console.WriteLine($"Password verification result for '{passwordToTest}': {isPasswordCorrect}");

        // If the password is correct, load the workbook with the password
        if (isPasswordCorrect)
        {
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Password = passwordToTest;

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine($"Workbook loaded. IsEncrypted: {workbook.Settings.IsEncrypted}");

            // Example operation: save a copy (save rule)
            workbook.Save("decrypted_copy.xlsx");
        }
    }
}