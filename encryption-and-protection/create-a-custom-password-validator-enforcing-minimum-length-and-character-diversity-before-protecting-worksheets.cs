// Title: Implement a custom password validator and protect an Excel worksheet with Aspose.Cells for .NET
// Common Searches: how to add custom password complexity check before worksheet protection using Aspose.Cells C# | Aspose.Cells protect worksheet with password that meets length and character requirements | C# example for validating Excel sheet password before calling Worksheet.Protect | enforce uppercase, lowercase, digit, special character rules for Excel worksheet protection in .NET | validate password strength prior to saving protected workbook with Aspose.Cells
// Tags: custom password validation Aspose.Cells .NET | worksheet protection with password criteria C# | Aspose.Cells Protect method password validation | Excel worksheet password strength enforcement | C# password complexity rules for Excel protection

using System;
using Aspose.Cells;

namespace AsposeCellsPasswordValidator
{
    // The sample creates a new workbook, defines a worksheet, validates a password for minimum length and required character types (uppercase, lowercase, digit, special character), protects the worksheet using Aspose.Cells with the validated password, and saves the workbook as ProtectedWorkbook.xlsx.
    class Program
    {
        // Minimum password length
        private const int MinLength = 8;

        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Example password to protect the worksheet
                string password = "P@ssw0rd!";

                // Validate password before applying protection
                if (!ValidatePassword(password))
                {
                    Console.WriteLine("Password does not meet the required criteria.");
                    return;
                }

                // Protect the worksheet with the validated password
                // The overload requires oldPassword; pass null when there is none
                sheet.Protect(ProtectionType.All, password, null);

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ProtectedWorkbook.xlsx");
                Console.WriteLine("Workbook saved with protected worksheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Custom password validator enforcing minimum length and character diversity
        private static bool ValidatePassword(string password)
        {
            if (string.IsNullOrEmpty(password) || password.Length < MinLength)
                return false;

            bool hasUpper = false, hasLower = false, hasDigit = false, hasSpecial = false;

            foreach (char c in password)
            {
                if (char.IsUpper(c)) hasUpper = true;
                else if (char.IsLower(c)) hasLower = true;
                else if (char.IsDigit(c)) hasDigit = true;
                else if (!char.IsLetterOrDigit(c)) hasSpecial = true;
            }

            // Require at least one uppercase, one lowercase, one digit, and one special character
            return hasUpper && hasLower && hasDigit && hasSpecial;
        }
    }
}
