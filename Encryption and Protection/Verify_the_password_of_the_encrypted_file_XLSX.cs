using System;
using System.IO;
using Aspose.Cells;

class VerifyEncryptedFilePassword
{
    static void Main()
    {
        // Path to the encrypted XLSX file
        string filePath = "encrypted.xlsx";

        // Password to verify
        string password = "test";

        // Open the file as a read‑only stream
        using (Stream stream = File.OpenRead(filePath))
        {
            // Use Aspose.Cells FileFormatUtil to check if the password is correct
            bool isValid = FileFormatUtil.VerifyPassword(stream, password);

            // Output the verification result
            Console.WriteLine($"Password '{password}' is valid: {isValid}");
        }
    }
}