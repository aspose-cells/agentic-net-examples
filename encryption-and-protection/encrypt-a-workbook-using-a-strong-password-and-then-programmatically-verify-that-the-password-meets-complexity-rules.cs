// Title: C# – Encrypt an Aspose.Cells Workbook with a Strong Password and Enforce Complexity Rules
// Description: This example shows how to create a new Workbook, validate that a password meets minimum‑length, uppercase, lowercase, digit and special‑character requirements, apply 256‑bit StrongCryptographicProvider encryption via SetEncryptionOptions, save the file, confirm the IsEncrypted flag, verify the password with FileFormatUtil.VerifyPassword, and finally open the protected workbook using LoadOptions.
// Keywords: Aspose.Cells | C# encryption | Excel password protection | strong password | password complexity validation | SetEncryptionOptions | EncryptionType.StrongCryptographicProvider | 256-bit encryption | FileFormatUtil.VerifyPassword | LoadOptions password | Workbook.Settings.Password
// Common Searches: How to encrypt an Excel workbook with a strong password using Aspose.Cells for .NET | C# code to check password complexity before applying Aspose.Cells workbook encryption | Verify password of an encrypted workbook programmatically with Aspose.Cells | Load a password‑protected workbook in Aspose.Cells using LoadOptions | Set 256‑bit encryption for an Aspose.Cells workbook
// Developer Intent: Secure a workbook with strong encryption and ensure the password complies with complexity policies before saving.
// Use Cases: Protect confidential Excel reports before distribution by applying 256‑bit encryption. | Enforce corporate password policies automatically when setting Workbook.Settings.Password. | Programmatically confirm that an encrypted file can be opened only with the correct password using FileFormatUtil.VerifyPassword. | Load a password‑protected workbook in a downstream process without manual intervention.
// AI Prompts: Generate C# code that encrypts an Aspose.Cells workbook with a 256‑bit password after validating the password contains at least 8 characters, an uppercase letter, a lowercase letter, a digit, and a special character. | Show how to use Aspose.Cells FileFormatUtil to verify a workbook's password and then open the file with LoadOptions in C#. | Write a C# function that checks password complexity according to NIST guidelines and applies the password to Workbook.Settings.Password with strong encryption.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordEncryptionDemo
{
    // This example shows how to create a new Workbook, validate that a password meets minimum‑length, uppercase, lowercase, digit and special‑character requirements, apply 256‑bit StrongCryptographicProvider encryption via SetEncryptionOptions, save the file, confirm the IsEncrypted flag, verify the password with FileFormatUtil.VerifyPassword, and finally open the protected workbook using LoadOptions.
    class Program
    {
        // Checks password complexity: minimum 8 chars, at least one upper, lower, digit and special character
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
            // Define a strong password
            string strongPassword = "Str0ngP@ssw0rd!";

            // Verify password meets complexity rules before applying
            if (!IsPasswordComplex(strongPassword))
            {
                Console.WriteLine("Password does not meet complexity requirements.");
                return;
            }

            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Encrypted Workbook Example");

            // Set workbook encryption password (WorkbookSettings.Password)
            workbook.Settings.Password = strongPassword;

            // Apply strong encryption options (SetEncryptionOptions)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the encrypted workbook (lifecycle rule)
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved with encryption to '{encryptedPath}'.");

            // Verify that the workbook is marked as encrypted
            Console.WriteLine($"IsEncrypted property after save: {workbook.Settings.IsEncrypted}");

            // Verify password correctness using FileFormatUtil (optional but demonstrates validation)
            using (Stream stream = File.OpenRead(encryptedPath))
            {
                bool isPasswordValid = FileFormatUtil.VerifyPassword(stream, strongPassword);
                Console.WriteLine($"Password validation via FileFormatUtil: {isPasswordValid}");
            }

            // Load the encrypted workbook using LoadOptions with the password
            LoadOptions loadOptions = new LoadOptions { Password = strongPassword };
            Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);
            Console.WriteLine("Workbook loaded successfully with the provided password.");

            // Confirm that the loaded workbook is still encrypted flag
            Console.WriteLine($"Loaded workbook IsEncrypted: {loadedWorkbook.Settings.IsEncrypted}");
        }
    }
}
