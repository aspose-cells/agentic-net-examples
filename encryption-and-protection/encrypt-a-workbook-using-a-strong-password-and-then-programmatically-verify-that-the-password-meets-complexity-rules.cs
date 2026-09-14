// Title: Encrypt an Excel workbook with a strong password and enforce custom password complexity using Aspose.Cells for .NET
// AI Prompts: Create C# code that checks a password for minimum length, uppercase, lowercase, digit, and special character requirements before calling Workbook.Protect to encrypt an .xlsx file with Aspose.Cells. | Refactor the example to raise a custom PasswordComplexityException instead of printing validation errors when the password does not meet the rules. | Update the sample to retrieve the password from an environment variable, validate it, and then protect the workbook with Aspose.Cells.
// Common Searches: asp.net encrypt excel file with password using Aspose.Cells and custom complexity validation | c# validate password strength before calling Workbook.Protect in Aspose.Cells | how to enforce password rules for Excel workbook encryption with Aspose.Cells .NET | example of protecting .xlsx with Aspose.Cells after password complexity check | Aspose.Cells workbook.Protect usage with password policy in C#
// Tags: encrypt workbook with Aspose.Cells Protect | password complexity validation C# Aspose.Cells | protect .xlsx file using Workbook.Protect | custom password policy for Excel encryption | Aspose.Cells workbook encryption example

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

// The program validates a password against length, uppercase, lowercase, digit, and special‑character rules, then uses Aspose.Cells Workbook.Protect to encrypt and save an .xlsx workbook with the verified strong password.
class Program
{
    static void Main()
    {
        try
        {
            // Define the password to be used for encryption
            string password = "Str0ng!Passw0rd";

            // Verify that the password meets complexity requirements
            if (!ValidatePasswordComplexity(password, out string validationError))
            {
                Console.WriteLine("Password validation failed: " + validationError);
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // (Optional) Add some sample data
            workbook.Worksheets[0].Cells["A1"].PutValue("Hello Aspose.Cells");

            // Protect the workbook with the validated password (strong encryption is applied automatically for .xlsx)
            // Note: Protect method expects ProtectionType first, then password
            workbook.Protect(ProtectionType.All, password);

            // Save the encrypted workbook
            string outputPath = "EncryptedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook encrypted and saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    // Checks password complexity: minimum 8 chars, at least one upper, lower, digit, and special character
    static bool ValidatePasswordComplexity(string password, out string errorMessage)
    {
        List<string> errors = new List<string>();

        if (password.Length < 8)
            errors.Add("minimum 8 characters");
        if (!Regex.IsMatch(password, "[A-Z]"))
            errors.Add("an uppercase letter");
        if (!Regex.IsMatch(password, "[a-z]"))
            errors.Add("a lowercase letter");
        if (!Regex.IsMatch(password, "[0-9]"))
            errors.Add("a digit");
        if (!Regex.IsMatch(password, "[^a-zA-Z0-9]"))
            errors.Add("a special character");

        if (errors.Count == 0)
        {
            errorMessage = string.Empty;
            return true;
        }

        errorMessage = "Password must contain " + string.Join(", ", errors) + ".";
        return false;
    }
}
