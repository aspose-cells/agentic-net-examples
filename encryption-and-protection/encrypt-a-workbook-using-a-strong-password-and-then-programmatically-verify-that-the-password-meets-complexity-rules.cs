// Title: Encrypt an Excel workbook with a strong password and enforce password complexity using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells encrypt workbook with a strong password and validate complexity | How to verify Excel file encryption by opening without password using Aspose.Cells | Load an encrypted XLSX file with password using Aspose.Cells LoadOptions in .NET | Set workbook structure protection with password in Aspose.Cells C# | Implement password complexity rules before protecting an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells workbook password protection | C# strong password validation for Excel | Aspose.Cells structure protection API | LoadOptions password-based workbook loading | check Excel file encryption status Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example defines a method that checks a password for minimum length, uppercase, lowercase, digit, and special‑character requirements. After validation, it creates a Workbook, assigns the strong password to Settings.Password, protects the workbook structure, and saves the file as an encrypted XLSX. It then demonstrates verification by attempting to open the file without a password (which fails) and successfully opening it with the correct password using LoadOptions.
class WorkbookEncryptionDemo
{
    // Checks if the password meets complexity requirements:
    // - Minimum 8 characters
    // - At least one uppercase letter
    // - At least one lowercase letter
    // - At least one digit
    // - At least one special character
    static bool IsPasswordComplex(string password)
    {
        if (string.IsNullOrEmpty(password) || password.Length < 8)
            return false;

        bool hasUpper = password.Any(char.IsUpper);
        bool hasLower = password.Any(char.IsLower);
        bool hasDigit = password.Any(char.IsDigit);
        bool hasSpecial = password.Any(ch => !char.IsLetterOrDigit(ch));

        return hasUpper && hasLower && hasDigit && hasSpecial;
    }

    static void Main()
    {
        // Define a strong password
        string strongPassword = "Str0ng!Passw0rd";

        // Verify password complexity before applying it
        if (!IsPasswordComplex(strongPassword))
        {
            Console.WriteLine("Password does not meet complexity requirements.");
            return;
        }

        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // Add some sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");

            // ---------- Encrypt the workbook with the strong password ----------
            // Set the password for opening the file (encryption)
            workbook.Settings.Password = strongPassword;

            // Optional: protect workbook structure
            workbook.Protect(ProtectionType.Structure, strongPassword);

            // Save the encrypted workbook
            string encryptedPath = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved and encrypted at: {Path.GetFullPath(encryptedPath)}");

            // ---------- Verify that the workbook is encrypted ----------
            if (File.Exists(encryptedPath))
            {
                // Attempt to load without a password (should fail)
                try
                {
                    Workbook loadFail = new Workbook(encryptedPath);
                    Console.WriteLine("Error: Workbook opened without password (encryption not applied).");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Successfully blocked opening without password: " + ex.Message);
                }

                // Load with the correct password
                try
                {
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = strongPassword
                    };
                    Workbook loadedWorkbook = new Workbook(encryptedPath, loadOptions);
                    Console.WriteLine("Workbook opened successfully with the correct password.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to open workbook with password: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Encrypted file not found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
