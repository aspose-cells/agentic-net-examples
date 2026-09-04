// Title: Implement a custom password validator and protect Excel worksheets using Aspose.Cells for .NET
// AI Prompts: Write a C# method that checks a password for minimum length, uppercase, lowercase, digit, and special character, then use it to call Worksheet.Protect(ProtectionType.All, password, string.Empty). | Refactor the example to loop through every worksheet in a workbook, applying the same validated password protection, and expose the complexity rules (min length, required character sets) as configurable parameters. | Show how to catch the ArgumentException thrown by the validator and log a clear error message before attempting to protect the worksheet.
// Common Searches: c# Aspose.Cells enforce password complexity before worksheet protection | how to validate Excel sheet password strength using Aspose.Cells .NET | protect multiple worksheets with a strong password in Aspose.Cells | custom password policy implementation for Aspose.Cells workbook protection | Aspose.Cells Protect method with custom validator example
// Tags: custom password validator Aspose.Cells | worksheet protection with password policy .NET | Aspose.Cells Protect method password enforcement | Excel workbook password complexity .NET | C# password strength check for Excel sheets

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The sample creates or loads an Excel workbook, validates each password using a custom method that enforces minimum length, uppercase, lowercase, digit, and special character requirements, then protects the worksheet(s) with Aspose.Cells' Protect method (ProtectionType.All) and saves the protected files.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook (create rule) ----------
            var newWorkbook = new Workbook(); // create rule
            var newSheet = newWorkbook.Worksheets[0];
            string newPassword = "Str0ng!Pass";

            // Validate and protect the new worksheet
            ValidatePassword(newPassword);
            // Protect the worksheet with a password (all protection types)
            // The third parameter is the old password; use empty string for a new protection
            newSheet.Protect(ProtectionType.All, newPassword, string.Empty);

            // Save the newly created and protected workbook (save rule)
            newWorkbook.Save("NewProtectedWorkbook.xlsx"); // save rule
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating new workbook: {ex.Message}");
        }

        try
        {
            // Ensure the template file exists; create a simple one if missing
            const string templatePath = "Template.xlsx";
            if (!File.Exists(templatePath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "TemplateSheet";
                tempWb.Save(templatePath);
            }

            // ---------- Load an existing workbook (load rule) ----------
            var loadedWorkbook = new Workbook(templatePath); // load rule
            var loadedSheet = loadedWorkbook.Worksheets[0];
            string loadedPassword = "An0ther#Pass";

            // Validate and protect the loaded worksheet
            ValidatePassword(loadedPassword);
            // Protect the worksheet with a password (all protection types)
            loadedSheet.Protect(ProtectionType.All, loadedPassword, string.Empty);

            // Save the loaded and protected workbook (save rule)
            loadedWorkbook.Save("LoadedProtectedWorkbook.xlsx"); // save rule
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing loaded workbook: {ex.Message}");
        }
    }

    // Custom password validator enforcing minimum length and character diversity
    static void ValidatePassword(string password)
    {
        const int minLength = 8;
        if (string.IsNullOrEmpty(password) || password.Length < minLength)
            throw new ArgumentException($"Password must be at least {minLength} characters long.");

        // At least one uppercase letter
        if (!Regex.IsMatch(password, @"[A-Z]"))
            throw new ArgumentException("Password must contain at least one uppercase letter.");

        // At least one lowercase letter
        if (!Regex.IsMatch(password, @"[a-z]"))
            throw new ArgumentException("Password must contain at least one lowercase letter.");

        // At least one digit
        if (!Regex.IsMatch(password, @"\d"))
            throw new ArgumentException("Password must contain at least one digit.");

        // At least one special character (non-word character)
        if (!Regex.IsMatch(password, @"[\W_]"))
            throw new ArgumentException("Password must contain at least one special character.");
    }
}
