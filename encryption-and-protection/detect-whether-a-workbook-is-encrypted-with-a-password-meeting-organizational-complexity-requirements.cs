// Title: C# – Detect Excel Workbook Encryption and Enforce Password Complexity with Aspose.Cells
// Description: A concise C# example that uses Aspose.Cells to (1) detect whether an Excel file is encrypted via FileFormatUtil, (2) validate the opening password against corporate complexity rules (minimum 8 characters, upper‑case, lower‑case, digit, special character), and (3) load the workbook only when the password complies, confirming the encryption flag through Workbook.Settings.
// Keywords: Aspose.Cells | C# Excel encryption detection | FileFormatUtil IsEncrypted | LoadOptions password | Workbook.Settings.IsEncrypted | Excel password complexity validation | password policy .NET | detect encrypted workbook without loading | secure Excel processing | organizational password rules
// Common Searches: how to check if an Excel file is password protected using Aspose.Cells | C# verify password meets complexity before opening encrypted workbook | FileFormatUtil detect encryption status without loading workbook | Aspose.Cells load encrypted workbook with password | validate Excel password policy in .NET
// Developer Intent: Identify the encryption state of an Excel workbook and ensure the provided password satisfies defined complexity criteria before attempting to open it.
// Use Cases: Quickly determine whether an incoming Excel file is encrypted to decide processing flow. | Enforce corporate password policies on encrypted workbooks prior to decryption. | Prevent unnecessary loading of large files when they are not password‑protected. | Log or audit encryption status and password compliance for security reviews.
// AI Prompts: Generate C# code that uses Aspose.Cells to return a boolean indicating if an Excel file is encrypted. | Create a method that checks a password against rules (≥8 chars, upper, lower, digit, special) before opening an encrypted workbook with Aspose.Cells. | Show how to handle a failed password‑complexity check when trying to load an encrypted Excel file using Aspose.Cells.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsPasswordComplexityCheck
{
    // A concise C# example that uses Aspose.Cells to (1) detect whether an Excel file is encrypted via FileFormatUtil, (2) validate the opening password against corporate complexity rules (minimum 8 characters, upper‑case, lower‑case, digit, special character), and (3) load the workbook only when the password complies, confirming the encryption flag through Workbook.Settings.
    class Program
    {
        // Define organizational password complexity rules
        // Minimum 8 characters, at least one uppercase, one lowercase, one digit, and one special character
        static bool IsPasswordComplex(string password)
        {
            if (string.IsNullOrEmpty(password) || password.Length < 8)
                return false;

            bool hasUpper = Regex.IsMatch(password, @"[A-Z]");
            bool hasLower = Regex.IsMatch(password, @"[a-z]");
            bool hasDigit = Regex.IsMatch(password, @"\d");
            bool hasSpecial = Regex.IsMatch(password, @"[\W_]");

            return hasUpper && hasLower && hasDigit && hasSpecial;
        }

        static void Main(string[] args)
        {
            // Path to the workbook to be examined
            string workbookPath = "sample.xlsx";

            // Password that is expected to open the workbook (if encrypted)
            // In a real scenario this would be obtained from a secure source
            string password = "P@ssw0rd!";

            // Detect file format and encryption status without loading the workbook
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(workbookPath);
            bool isEncrypted = formatInfo.IsEncrypted;

            Console.WriteLine($"Workbook encrypted: {isEncrypted}");

            if (isEncrypted)
            {
                // Verify password complexity according to organizational policy
                bool passwordComplex = IsPasswordComplex(password);
                Console.WriteLine($"Provided password meets complexity requirements: {passwordComplex}");

                if (passwordComplex)
                {
                    // Attempt to load the workbook using the provided password
                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = password
                    };

                    Workbook workbook = new Workbook(workbookPath, loadOptions);

                    // After successful load, double‑check encryption flag via settings
                    bool settingsEncrypted = workbook.Settings.IsEncrypted;
                    Console.WriteLine($"Workbook.Settings.IsEncrypted after load: {settingsEncrypted}");

                    // Dispose workbook when done
                    workbook.Dispose();
                }
                else
                {
                    Console.WriteLine("Password does not satisfy complexity rules; cannot load workbook.");
                }
            }
            else
            {
                Console.WriteLine("Workbook is not encrypted; no password validation required.");
            }
        }
    }
}
