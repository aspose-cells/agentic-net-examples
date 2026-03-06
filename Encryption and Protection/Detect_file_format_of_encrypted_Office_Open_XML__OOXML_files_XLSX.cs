using System;
using System.IO;
using Aspose.Cells;

class DetectEncryptedOOXML
{
    static void Main()
    {
        // Path to the encrypted OOXML file (e.g., .xlsx)
        string filePath = "encrypted.xlsx";

        // Password used to protect the file
        string password = "test";

        // Detect the file format using the password overload
        using (FileStream stream = File.OpenRead(filePath))
        {
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(stream, password);
            Console.WriteLine($"Detected File Format: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
        }

        // Verify whether the supplied password is correct
        using (FileStream stream = File.OpenRead(filePath))
        {
            bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
            Console.WriteLine($"Password valid: {isPasswordValid}");
        }
    }
}