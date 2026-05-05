using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the encrypted XLSX file (OOXML spreadsheet)
        string filePath = "encrypted.xlsx";

        // Password used to protect the workbook
        string password = "test";

        // Detect the file format and encryption status using the password
        using (FileStream stream = File.OpenRead(filePath))
        {
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(stream, password);
            Console.WriteLine($"Detected File Format Type: {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
        }

        // Optional: verify that the supplied password is correct
        using (FileStream stream = File.OpenRead(filePath))
        {
            bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
            Console.WriteLine($"Password valid: {isPasswordValid}");
        }
    }
}