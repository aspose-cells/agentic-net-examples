// Title: Encrypt an Excel workbook with AES‑256 and enforce password complexity using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, validate that a password meets minimum length, uppercase, lowercase, digit and special‑character requirements, apply AES‑256 encryption via SetEncryptionOptions, save the file, and programmatically confirm the password using LoadOptions and FileFormatUtil.VerifyPassword.
// Keywords: Aspose.Cells encryption C# | AES 256 Excel password | password complexity validation .NET | Workbook.SetEncryptionOptions | FileFormatUtil VerifyPassword | secure Excel file Aspose | LoadOptions password Excel
// Common Searches: Aspose.Cells encrypt Excel with AES‑256 | C# check password strength before workbook encryption | verify Excel file password programmatically Aspose | how to set strong encryption for .xlsx using Aspose.Cells | validate password rules Aspose.Cells .NET
// Developer Intent: Secure an Excel workbook with AES‑256 encryption while ensuring the password satisfies defined complexity rules before saving.
// Use Cases: Generate a new workbook, run a complexity check, and save it encrypted with a strong password. | Reload the encrypted file with LoadOptions to confirm decryption works with the correct password. | Use FileFormatUtil.VerifyPassword on the saved stream to programmatically validate the password.
// AI Prompts: Write C# code that encrypts an Aspose.Cells workbook with AES‑256 and enforces a password containing at least 8 characters, an uppercase letter, a lowercase letter, a digit, and a special character. | Provide a reusable method for password‑complexity validation that integrates with Aspose.Cells encryption workflow. | Show how to programmatically verify an encrypted workbook’s password using LoadOptions and FileFormatUtil in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    // Demonstrates how to create a workbook, validate that a password meets minimum length, uppercase, lowercase, digit and special‑character requirements, apply AES‑256 encryption via SetEncryptionOptions, save the file, and programmatically confirm the password using LoadOptions and FileFormatUtil.VerifyPassword.
    class Program
    {
        // Checks password complexity: minimum 8 chars, at least one upper, lower, digit, special character
        static bool IsPasswordComplex(string password)
        {
            if (string.IsNullOrEmpty(password) || password.Length < 8)
                return false;

            bool hasUpper = false, hasLower = false, hasDigit = false, hasSpecial = false;
            foreach (char c in password)
            {
                if (char.IsUpper(c)) hasUpper = true;
                else if (char.IsLower(c)) hasLower = true;
                else if (char.IsDigit(c)) hasDigit = true;
                else hasSpecial = true;
            }
            return hasUpper && hasLower && hasDigit && hasSpecial;
        }

        static void Main()
        {
            // Define a strong password that satisfies the complexity rules
            string strongPassword = "Str0ng!Pass";

            // Verify password complexity before applying it
            if (!IsPasswordComplex(strongPassword))
            {
                Console.WriteLine("Password does not meet complexity requirements.");
                return;
            }

            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted content");

            // Apply encryption password
            wb.Settings.Password = strongPassword;

            // Set strong encryption options (AES 256)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the encrypted workbook
            string encryptedPath = "EncryptedWorkbook.xlsx";
            wb.Save(encryptedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved with encryption to '{encryptedPath}'.");

            // Verify that the workbook is marked as encrypted
            Console.WriteLine($"IsEncrypted property after save: {wb.Settings.IsEncrypted}");

            // Load the encrypted workbook using the correct password
            LoadOptions loadOptions = new LoadOptions { Password = strongPassword };
            Workbook loadedWb = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine("Workbook loaded successfully with correct password.");

            // Verify that the loaded workbook is indeed encrypted
            Console.WriteLine($"IsEncrypted after load: {loadedWb.Settings.IsEncrypted}");

            // Additional verification using FileFormatUtil
            using (Stream stream = File.OpenRead(encryptedPath))
            {
                bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, strongPassword);
                Console.WriteLine($"FileFormatUtil password verification result: {isPasswordValid}");
            }
        }
    }
}
