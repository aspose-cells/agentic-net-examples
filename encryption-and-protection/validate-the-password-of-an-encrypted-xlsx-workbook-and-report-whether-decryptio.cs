using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ValidateEncryptedWorkbookPassword
    {
        public static void Run()
        {
            // Path to the encrypted workbook
            string filePath = "encrypted.xlsx";

            // Passwords to test
            string correctPassword = "test";
            string wrongPassword = "1234";

            // Verify correct password
            using (Stream stream = File.OpenRead(filePath))
            {
                bool isValid = FileFormatUtil.VerifyPassword(stream, correctPassword);
                Console.WriteLine($"Password '{correctPassword}' is valid: {isValid}");
            }

            // Verify wrong password
            using (Stream stream = File.OpenRead(filePath))
            {
                bool isValid = FileFormatUtil.VerifyPassword(stream, wrongPassword);
                Console.WriteLine($"Password '{wrongPassword}' is valid: {isValid}");
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            ValidateEncryptedWorkbookPassword.Run();
        }
    }
}