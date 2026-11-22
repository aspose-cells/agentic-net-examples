// Title: Check if an Excel workbook is password‑protected and validate password complexity with Aspose.Cells for .NET
// Common Searches: how to determine if an Excel file is encrypted using Aspose.Cells C# | validate Excel workbook password against complexity policy with Aspose.Cells | catch CellsException for wrong password when opening a protected workbook in .NET | use LoadOptions to open a password‑protected XLSX and test password strength in C# | programmatically verify Excel encryption and password rules with Aspose.Cells
// Tags: Aspose.Cells LoadOptions password protection | detect encrypted XLSX workbook Aspose.Cells | validate password complexity C# regex | handle CellsException incorrect password | programmatic Excel encryption check .NET

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an XLSX file with Aspose.Cells using LoadOptions and a supplied password, catches CellsException to identify encrypted workbooks with an incorrect password, and runs a regex‑based method to ensure the password complies with organizational complexity requirements (minimum length, uppercase, lowercase, digit, special character).
class Program
{
    static void Main()
    {
        // Path to the workbook to check
        string filePath = "encrypted.xlsx";

        // Password to attempt opening the workbook (if known)
        string password = "YourPassword123!";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // LoadOptions with the password for encrypted workbooks
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = password
        };

        Workbook workbook = null;

        try
        {
            // Load the workbook using the provided password
            workbook = new Workbook(filePath, loadOptions);

            // If loading succeeds, the password is correct (or the file is not encrypted)
            bool meetsComplexity = IsPasswordComplex(password);
            Console.WriteLine($"Workbook loaded successfully. Password meets complexity requirements: {meetsComplexity}");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for incorrect passwords
            if (ex.Message != null && ex.Message.IndexOf("incorrect password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Workbook is encrypted, but the provided password is incorrect.");
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Other errors (e.g., file corrupted, unsupported format)
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }

    // Organizational password complexity rule:
    // Minimum 8 characters, at least one uppercase, one lowercase, one digit, and one special character
    static bool IsPasswordComplex(string pwd)
    {
        if (string.IsNullOrEmpty(pwd) || pwd.Length < 8)
            return false;

        bool hasUpper = Regex.IsMatch(pwd, "[A-Z]");
        bool hasLower = Regex.IsMatch(pwd, "[a-z]");
        bool hasDigit = Regex.IsMatch(pwd, "[0-9]");
        bool hasSpecial = Regex.IsMatch(pwd, "[^a-zA-Z0-9]");

        return hasUpper && hasLower && hasDigit && hasSpecial;
    }
}
