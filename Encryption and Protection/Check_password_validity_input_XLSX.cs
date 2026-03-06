using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be checked
            string filePath = "input.xlsx";

            // Password to validate
            string passwordToCheck = "testPassword";

            // Verify that the file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Open the file as a read‑only stream
            using (Stream stream = File.OpenRead(filePath))
            {
                // Use Aspose.Cells FileFormatUtil to verify the password for encrypted OOXML files
                bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, passwordToCheck);

                // Output the result
                Console.WriteLine($"Password '{passwordToCheck}' is valid: {isPasswordValid}");
            }
        }
    }
}