// Title: C# Custom Password Validator for Aspose.Cells Worksheet Protection (Min Length & Complexity)
// Description: Demonstrates how to enforce a minimum 8‑character length and require at least three of four character groups (uppercase, lowercase, digit, special) using regular expressions. The validator runs before assigning the password to the worksheet's Protection object, then the sheet is protected with ProtectionType.All, the password is verified with VerifyPassword, and the workbook is saved as an Excel file.
// Keywords: Aspose.Cells password validation | C# worksheet protection | Excel password complexity | custom password validator .NET | Protect worksheet Aspose.Cells | VerifyPassword C# | regex password rules | Excel security Aspose
// Common Searches: how to enforce password complexity with Aspose.Cells C# | validate Excel worksheet password before Protect method | Aspose.Cells verify worksheet protection password | custom password rules for Excel sheet protection .NET | C# regex password validator for Aspose.Cells
// Developer Intent: Ensure that the password used for worksheet protection meets defined length and character‑type requirements before applying protection with Aspose.Cells.
// Use Cases: Validate user‑entered passwords against strength rules before calling sheet.Protect to avoid weak protection. | Apply a strong, pre‑validated password while configuring granular worksheet protection options (allow/deny actions). | Programmatically confirm that the saved workbook’s protection password satisfies the complexity policy using VerifyPassword.
// AI Prompts: Create a reusable C# method that checks password strength (minimum 8 characters, at least three of uppercase, lowercase, digit, special) for Aspose.Cells worksheet protection. | Show how to integrate the custom password validator into an existing Aspose.Cells workflow and return a clear error message when validation fails. | Write unit tests in C# for the IsValidPassword method covering all combinations of character categories and edge cases.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsPasswordValidator
{
    // Demonstrates how to enforce a minimum 8‑character length and require at least three of four character groups (uppercase, lowercase, digit, special) using regular expressions. The validator runs before assigning the password to the worksheet's Protection object, then the sheet is protected with ProtectionType.All, the password is verified with VerifyPassword, and the workbook is saved as an Excel file.
    class Program
    {
        // Validates that the password meets minimum length and character diversity requirements.
        // Minimum length: 8 characters.
        // Must contain at least three of the following categories:
        //   - Uppercase letters
        //   - Lowercase letters
        //   - Digits
        //   - Special characters
        static bool IsValidPassword(string password)
        {
            if (string.IsNullOrEmpty(password) || password.Length < 8)
                return false;

            int categories = 0;
            if (Regex.IsMatch(password, @"[A-Z]")) categories++; // Uppercase
            if (Regex.IsMatch(password, @"[a-z]")) categories++; // Lowercase
            if (Regex.IsMatch(password, @"\d")) categories++;    // Digit
            if (Regex.IsMatch(password, @"[\W_]")) categories++; // Special

            return categories >= 3;
        }

        static void Main(string[] args)
        {
            // Example password to be used for worksheet protection
            string password = "Str0ng!Pass";

            // Validate password before applying protection
            if (!IsValidPassword(password))
            {
                Console.WriteLine("Password does not meet the required complexity rules.");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure worksheet protection options
            Protection protection = sheet.Protection;
            protection.AllowDeletingColumn = false;
            protection.AllowDeletingRow = false;
            protection.AllowFormattingCell = true;
            protection.AllowInsertingRow = true;
            protection.AllowSelectingLockedCell = true;
            protection.Password = password; // Set the validated password

            // Protect the worksheet with all protection types using the password
            sheet.Protect(ProtectionType.All, password, null);

            // Verify that the password works using VerifyPassword
            bool isCorrect = sheet.Protection.VerifyPassword(password);
            Console.WriteLine($"Password verification result: {isCorrect}");

            // Save the workbook
            workbook.Save("ProtectedWorksheet.xlsx");
        }
    }
}
