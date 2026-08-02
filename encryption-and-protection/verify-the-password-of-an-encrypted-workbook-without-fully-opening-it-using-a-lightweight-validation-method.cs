using System;
using System.IO;
using Aspose.Cells;

class VerifyWorkbookPassword
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        // Correct password to test
        string correctPassword = "test";

        // Verify the correct password without fully loading the workbook
        using (Stream stream = File.OpenRead(filePath))
        {
            bool isValid = FileFormatUtil.VerifyPassword(stream, correctPassword);
            Console.WriteLine($"Password '{correctPassword}' is valid: {isValid}");
        }

        // Verify an incorrect password using the same lightweight method
        using (Stream stream = File.OpenRead(filePath))
        {
            bool isValid = FileFormatUtil.VerifyPassword(stream, "wrong");
            Console.WriteLine($"Password 'wrong' is valid: {isValid}");
        }
    }
}