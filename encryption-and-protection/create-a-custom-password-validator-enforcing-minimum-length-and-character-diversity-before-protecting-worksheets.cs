// Title: C# Custom Password Validator for Worksheet Protection with Aspose.Cells
// Description: Demonstrates how to enforce a minimum length and character‑type diversity (uppercase, lowercase, digit, special) in a password, then protect an Excel worksheet using Aspose.Cells .NET and save the workbook.
// Keywords: Aspose.Cells C# password validation | worksheet protection Aspose.Cells | custom password rules .NET | Excel sheet protect with password | password strength validator C# | Aspose.Cells ProtectionType.All | save protected workbook
// Common Searches: Aspose.Cells enforce password complexity before sheet protection | C# example custom password validator for Excel worksheet | How to protect an Aspose.Cells worksheet with a strong password | Validate password length and character categories in .NET | Aspose.Cells protect worksheet with specific options
// Developer Intent: Validate a password’s strength (length and character categories) before applying it to protect an Excel worksheet with Aspose.Cells.
// Use Cases: Check user‑entered passwords meet security policies before worksheet protection. | Apply full protection to a worksheet while allowing only cell formatting. | Save a workbook after successful password validation and protection.
// AI Prompts: Create a reusable C# method that validates password strength (min length, uppercase, lowercase, digit, special) for Aspose.Cells worksheet protection. | Show how to integrate the custom validator into a workflow that protects multiple worksheets with different passwords. | Provide error‑handling code that returns a clear message when the password fails validation before calling sheet.Protect.

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordValidator
{
    // Demonstrates how to enforce a minimum length and character‑type diversity (uppercase, lowercase, digit, special) in a password, then protect an Excel worksheet using Aspose.Cells .NET and save the workbook.
    class Program
    {
        // Minimum password length
        private const int MinLength = 8;

        // Validate password for length and character diversity
        private static bool IsValidPassword(string password)
        {
            if (string.IsNullOrEmpty(password) || password.Length < MinLength)
                return false;

            bool hasUpper = false, hasLower = false, hasDigit = false, hasSpecial = false;

            foreach (char c in password)
            {
                if (char.IsUpper(c)) hasUpper = true;
                else if (char.IsLower(c)) hasLower = true;
                else if (char.IsDigit(c)) hasDigit = true;
                else hasSpecial = true;
            }

            // Require at least three of the four categories
            int categories = (hasUpper ? 1 : 0) + (hasLower ? 1 : 0) + (hasDigit ? 1 : 0) + (hasSpecial ? 1 : 0);
            return categories >= 3;
        }

        static void Main()
        {
            // Example password to be used for worksheet protection
            string password = "Str0ng!Pass";

            // Validate password before applying protection
            if (!IsValidPassword(password))
            {
                Console.WriteLine("Password does not meet the required criteria.");
                return;
            }

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure worksheet protection options
            Protection protection = sheet.Protection;
            protection.AllowDeletingColumn = false;
            protection.AllowDeletingRow = false;
            protection.AllowFormattingCell = true;
            protection.AllowEditingContent = false;

            // Apply password and protect the worksheet (using overload with password)
            sheet.Protect(ProtectionType.All, password, null);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ProtectedWorksheet.xlsx");
        }
    }
}
