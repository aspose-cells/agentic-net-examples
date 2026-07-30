// Title: Custom Strong Password Validator for Worksheet Protection with Aspose.Cells (C#)
// Description: Implements a regex‑based IsPasswordStrong method, validates the password, configures worksheet protection options, applies ProtectionType.All, and saves a securely protected Excel file using Aspose.Cells.
// Keywords: Aspose.Cells C# worksheet protection | password complexity validator | regex password strength | Excel file encryption | custom password rules | ProtectionType.All | secure Excel workbook | Aspose.Cells API security | worksheet protect password | C# password validation example
// Common Searches: how to enforce password complexity with Aspose.Cells | c# example custom password validator for Excel sheet protection | protect worksheet using Aspose.Cells after password check | Aspose.Cells password strength validation code | set worksheet protection options C# Aspose
// Developer Intent: Validate a password against strength criteria before applying worksheet protection with Aspose.Cells.
// Use Cases: Enforce corporate password policies prior to protecting Excel worksheets. | Prevent weak passwords from being used with Worksheet.Protect in automated report generators. | Combine custom protection flags (disable delete/format) with a verified strong password.
// AI Prompts: Create a reusable C# method that checks password length, uppercase, lowercase, digit, and special character, then integrates with Aspose.Cells worksheet protection. | Write code to read a password from console input, repeatedly prompt until it passes the custom validator, and protect the worksheet. | Generate a snippet that logs a warning and aborts workbook saving when the password fails the strength test.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsCustomPasswordValidator
{
    // Implements a regex‑based IsPasswordStrong method, validates the password, configures worksheet protection options, applies ProtectionType.All, and saves a securely protected Excel file using Aspose.Cells.
    class Program
    {
        // Custom password validator enforcing minimum length and character diversity
        static bool IsPasswordStrong(string password, int minLength = 8)
        {
            if (string.IsNullOrEmpty(password) || password.Length < minLength)
                return false;

            // At least one uppercase, one lowercase, one digit, and one special character
            bool hasUpper = Regex.IsMatch(password, "[A-Z]");
            bool hasLower = Regex.IsMatch(password, "[a-z]");
            bool hasDigit = Regex.IsMatch(password, "[0-9]");
            bool hasSpecial = Regex.IsMatch(password, "[^a-zA-Z0-9]");

            return hasUpper && hasLower && hasDigit && hasSpecial;
        }

        static void Main(string[] args)
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Define the password to protect the worksheet
            string password = "StrongP@ssw0rd";

            // Validate password before applying protection
            if (!IsPasswordStrong(password))
            {
                Console.WriteLine("Password does not meet the required strength criteria.");
                return;
            }

            // Set protection options (optional)
            Protection protection = sheet.Protection;
            protection.AllowDeletingColumn = false;
            protection.AllowDeletingRow = false;
            protection.AllowFormattingCell = false;
            protection.AllowInsertingRow = false;
            protection.Password = password; // assign password

            // Protect the worksheet with all protection types
            sheet.Protect(ProtectionType.All, password, null);

            // Save the workbook (save rule)
            workbook.Save("ProtectedWorksheet.xlsx");

            Console.WriteLine("Worksheet protected successfully with a strong password.");
        }
    }
}
