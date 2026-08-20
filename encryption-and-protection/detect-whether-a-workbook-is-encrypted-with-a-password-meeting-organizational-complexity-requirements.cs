// Title: Identify encrypted Excel files and enforce password strength using Aspose.Cells (.NET)
// Description: A C# sample that uses Aspose.Cells to discover if an .xlsx workbook is password‑protected without opening it (FileFormatUtil.DetectFileFormat), open the file with a given password (LoadOptions), verify the workbook’s encryption flag, and assess the password against a typical corporate policy (minimum 8 characters, upper‑case, lower‑case, digit, special symbol).
// Keywords: Aspose.Cells C# encryption detection | Excel file password protection check | FileFormatUtil IsEncrypted | LoadOptions password Excel | validate Excel password policy | .NET workbook encryption | encrypted .xlsx detection | password complexity rule C# | Aspose.Cells security features | Excel workbook protection audit
// Common Searches: How to know if an Excel workbook is password protected with Aspose.Cells | C# code to open encrypted .xlsx using a supplied password | Check Excel file encryption status without loading the workbook | Validate that an Excel password meets corporate complexity requirements | Aspose.Cells example for encrypted workbook handling
// Developer Intent: Determine whether an Excel workbook is encrypted and confirm that the supplied password satisfies organizational complexity standards before further processing.
// Use Cases: Skip decryption attempts for files that are not password‑protected, saving resources. | Enforce company password policies automatically when opening protected workbooks. | Log encryption status and password‑policy compliance for audit and compliance reporting. | Provide user feedback on password strength before granting access to encrypted data.
// AI Prompts: Generate C# code with Aspose.Cells that checks if an .xlsx file is encrypted and then validates the password against a custom complexity rule. | Show how to catch and handle incorrect‑password exceptions when loading a protected workbook using Aspose.Cells. | Explain how FileFormatUtil.DetectFileFormat can be used to retrieve the IsEncrypted flag without opening the Excel file.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionCheck
{
    // A C# sample that uses Aspose.Cells to discover if an .xlsx workbook is password‑protected without opening it (FileFormatUtil.DetectFileFormat), open the file with a given password (LoadOptions), verify the workbook’s encryption flag, and assess the password against a typical corporate policy (minimum 8 characters, upper‑case, lower‑case, digit, special symbol).
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be examined
            string filePath = "sample.xlsx";

            // Password supplied by the user (could be obtained from UI, config, etc.)
            string suppliedPassword = "P@ssw0rd!";

            // Detect file format and encryption status without opening the file
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            Console.WriteLine($"File is encrypted: {formatInfo.IsEncrypted}");

            if (!formatInfo.IsEncrypted)
            {
                Console.WriteLine("Workbook is not encrypted. No password validation required.");
                return;
            }

            // Attempt to load the encrypted workbook using the supplied password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = suppliedPassword
            };

            try
            {
                // Load the workbook (uses the provided load rule)
                Workbook workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully with the supplied password.");

                // Verify that the workbook settings also report encryption
                Console.WriteLine($"Workbook.Settings.IsEncrypted: {workbook.Settings.IsEncrypted}");

                // Check password complexity according to organizational policy
                bool meetsComplexity = IsPasswordComplex(suppliedPassword);
                Console.WriteLine($"Password meets complexity requirements: {meetsComplexity}");
            }
            catch (Exception ex)
            {
                // Loading failed – likely due to an incorrect password
                Console.WriteLine($"Failed to open workbook. Reason: {ex.Message}");
            }
        }

        // Determines whether a password satisfies typical complexity rules:
        // Minimum 8 characters, at least one uppercase, one lowercase, one digit, and one special character.
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
                else if (!char.IsWhiteSpace(c)) hasSpecial = true;
            }

            return hasUpper && hasLower && hasDigit && hasSpecial;
        }
    }
}
