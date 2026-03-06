using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordVerification
{
    class Program
    {
        static void Main()
        {
            // Path to the encrypted Excel file
            string filePath = "encrypted.xlsx";

            // Passwords to test
            string correctPassword = "test";
            string wrongPassword = "1234";

            // Verify with the correct password
            using (Stream stream = File.OpenRead(filePath))
            {
                bool isValid = FileFormatUtil.VerifyPassword(stream, correctPassword);
                Console.WriteLine($"Password '{correctPassword}' is valid: {isValid}");
            }

            // Verify with an incorrect password
            using (Stream stream = File.OpenRead(filePath))
            {
                bool isValid = FileFormatUtil.VerifyPassword(stream, wrongPassword);
                Console.WriteLine($"Password '{wrongPassword}' is valid: {isValid}");
            }

            // Optional: check if the file is encrypted before attempting verification
            FileFormatInfo fileInfo = FileFormatUtil.DetectFileFormat(filePath);
            Console.WriteLine($"Is the file encrypted? {fileInfo.IsEncrypted}");
        }
    }
}