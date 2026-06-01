using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordValidation
{
    class Program
    {
        static void Main()
        {
            // Path to the encrypted workbook
            string filePath = "encrypted.xlsx";

            // Password to validate
            string passwordToTest = "test";

            // Validate the password using a lightweight method (does not fully load the workbook)
            bool isValid = VerifyWorkbookPassword(filePath, passwordToTest);

            Console.WriteLine($"Password '{passwordToTest}' is valid: {isValid}");
        }

        /// <summary>
        /// Verifies whether the supplied password can open the encrypted workbook.
        /// This method uses FileFormatUtil.VerifyPassword which checks the password
        /// without fully loading the workbook into memory.
        /// </summary>
        /// <param name="filePath">Full path to the encrypted workbook file.</param>
        /// <param name="password">Password to verify.</param>
        /// <returns>True if the password is correct; otherwise false.</returns>
        static bool VerifyWorkbookPassword(string filePath, string password)
        {
            // Open the file as a read‑only stream
            using (Stream stream = File.OpenRead(filePath))
            {
                // FileFormatUtil.VerifyPassword returns true if the password matches
                return FileFormatUtil.VerifyPassword(stream, password);
            }
        }
    }
}