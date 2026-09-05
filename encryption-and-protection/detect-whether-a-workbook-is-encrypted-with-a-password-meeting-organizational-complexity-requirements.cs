// Title: Determine if an Excel workbook is password‑protected and validate the password against complexity requirements with Aspose.Cells for .NET
// AI Prompts: Write a C# routine that uses Aspose.Cells to detect whether a given .xlsx file is encrypted and then checks the supplied password for length, uppercase, lowercase, digit, and special‑character criteria. | Create a method that returns true only when the workbook loads successfully with a password that satisfies organizational complexity rules, otherwise returns false. | Implement error handling to differentiate between unencrypted files, incorrect passwords, and passwords that fail complexity validation.
// Common Searches: Aspose.Cells C# how to detect encrypted Excel file | C# verify password complexity for protected workbook using Aspose | Check if .xlsx is password protected and meets policy with .NET | Load encrypted Excel with password and enforce complexity rules Aspose.Cells | Determine if workbook requires password and validate strength in C#
// Tags: detect encrypted workbook Aspose.Cells | validate password complexity C# | load protected XLSX with Aspose.Cells | encryption detection Excel .NET | password policy enforcement Aspose.Cells | workbook password verification C#

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example provides a WorkbookEncryptionChecker class that first attempts to open an Excel file with Aspose.Cells without a password to see if it is encrypted. If encryption is detected, it reloads the file using the supplied password via LoadOptions. Upon successful load, the password is evaluated against typical organizational complexity rules (minimum 8 characters, uppercase, lowercase, digit, special character). The method returns true only when the workbook is encrypted and the password meets all complexity criteria.
public class WorkbookEncryptionChecker
{
    // Checks if a password meets typical organizational complexity requirements:
    // - Minimum 8 characters
    // - At least one uppercase letter
    // - At least one lowercase letter
    // - At least one digit
    // - At least one special character
    private static bool IsPasswordComplex(string password)
    {
        if (string.IsNullOrEmpty(password) || password.Length < 8)
            return false;

        bool hasUpper = Regex.IsMatch(password, "[A-Z]");
        bool hasLower = Regex.IsMatch(password, "[a-z]");
        bool hasDigit = Regex.IsMatch(password, "[0-9]");
        bool hasSpecial = Regex.IsMatch(password, "[^a-zA-Z0-9]");

        return hasUpper && hasLower && hasDigit && hasSpecial;
    }

    // Determines whether the workbook is encrypted and whether the supplied password
    // satisfies the complexity rules.
    // Returns true only if the workbook is encrypted AND the password is complex.
    public static bool IsEncryptedWithComplexPassword(string workbookPath, string password)
    {
        // Ensure the file exists to avoid FileNotFoundException.
        if (!File.Exists(workbookPath))
            return false;

        // First, try loading without a password. If it succeeds, the workbook is not encrypted.
        try
        {
            var wb = new Workbook(workbookPath);
            // Loaded successfully → not encrypted.
            return false;
        }
        catch (Exception)
        {
            // Assume the failure is due to encryption; proceed to load with the supplied password.
        }

        // Attempt to load with the supplied password.
        try
        {
            var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            var wb = new Workbook(workbookPath, loadOptions);
            // If we reach here, the password was correct and the workbook is encrypted.
            return IsPasswordComplex(password);
        }
        catch (Exception)
        {
            // Loading failed (incorrect password or other issue) → does not meet requirements.
            return false;
        }
    }
}

// Simple entry point for demonstration/testing purposes.
public class Program
{
    public static void Main(string[] args)
    {
        // Example usage:
        // args[0] = path to workbook, args[1] = password to test.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <exe> <workbookPath> <password>");
            return;
        }

        string workbookPath = args[0];
        string password = args[1];

        try
        {
            bool result = WorkbookEncryptionChecker.IsEncryptedWithComplexPassword(workbookPath, password);
            Console.WriteLine($"Workbook encrypted with complex password: {result}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
