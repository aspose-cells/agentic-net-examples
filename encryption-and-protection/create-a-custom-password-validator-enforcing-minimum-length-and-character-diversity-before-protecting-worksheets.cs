// Title: Create a custom password strength validator and apply worksheet protection with Aspose.Cells for Excel files in C#
// Common Searches: how to enforce password complexity before protecting an Excel worksheet using Aspose.Cells C# | C# example for validating password length and required character types prior to workbook protection with Aspose.Cells | Aspose.Cells Protect method with custom password validator in .NET | validate password contains uppercase, lowercase, digit and special character before Excel file protection | sample code to protect all worksheets after password strength check using Aspose.Cells
// Tags: custom password validator Aspose.Cells C# | worksheet protection password Aspose.Cells | Excel workbook protection password complexity | Aspose.Cells Protect method usage | validate password strength before Excel protection

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, adds data, and defines a password. It runs a custom ValidatePassword method that checks for a minimum of eight characters, at least one uppercase letter, one lowercase letter, one digit, and one special character. If the password meets these criteria, each worksheet is secured with ws.Protect(ProtectionType.All, password, null). The protected workbook is saved to a specified path, with directory creation and exception handling included.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Password to protect the worksheets
            string password = "P@ssw0rd123";

            // Validate the password before applying protection
            if (!ValidatePassword(password))
            {
                Console.WriteLine("Password does not meet the required security criteria.");
                return;
            }

            // Protect each worksheet with the validated password
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // The third parameter is the old password; not needed here, so pass null
                ws.Protect(ProtectionType.All, password, null);
            }

            // Determine output path and ensure directory exists
            string outputPath = "ProtectedWorkbook.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Custom password validator:
    // - Minimum length: 8 characters
    // - At least one uppercase letter
    // - At least one lowercase letter
    // - At least one digit
    // - At least one special character
    static bool ValidatePassword(string pwd)
    {
        if (string.IsNullOrEmpty(pwd) || pwd.Length < 8)
            return false;

        bool hasUpper = false, hasLower = false, hasDigit = false, hasSpecial = false;

        foreach (char c in pwd)
        {
            if (char.IsUpper(c))
                hasUpper = true;
            else if (char.IsLower(c))
                hasLower = true;
            else if (char.IsDigit(c))
                hasDigit = true;
            else if (!char.IsWhiteSpace(c) && !char.IsLetterOrDigit(c))
                hasSpecial = true;
        }

        return hasUpper && hasLower && hasDigit && hasSpecial;
    }
}
