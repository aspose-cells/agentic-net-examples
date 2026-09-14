// Title: How to verify a worksheet's protection password in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# method that opens an .xlsx workbook, selects a worksheet by name, and uses Worksheet.Protection.VerifyPassword to return true when the supplied password matches the sheet's protection password. | Build a console application that calls the validator method to determine if a given password can unlock a protected worksheet and outputs the validation result.
// Common Searches: Aspose.Cells C# verify worksheet protection password programmatically | Check if an Excel sheet is password‑protected using Aspose.Cells .NET | C# code to validate worksheet password before editing with Aspose.Cells
// Tags: Aspose.Cells worksheet protection verification | C# Worksheet.Protection.VerifyPassword example | load Excel workbook and check sheet password Aspose.Cells | validate worksheet password .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetPasswordValidatorApp
{
    // Provides a static ValidatePassword method that loads a workbook, retrieves a worksheet by name, and calls Worksheet.Protection.VerifyPassword to determine if the supplied password matches the worksheet's protection password, returning false for missing files, missing sheets, or any exception.
    public class WorksheetPasswordValidator
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="sheetName">Name of the worksheet to check.</param>
        /// <param name="password">Password to validate.</param>
        /// <returns>True if the password matches the worksheet's protection password; otherwise false.</returns>
        public static bool ValidatePassword(string filePath, string sheetName, string password)
        {
            // Ensure the file exists to avoid FileNotFoundException.
            if (!File.Exists(filePath))
                return false;

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(filePath);

                // Try to get the worksheet by name.
                Worksheet worksheet = workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    // Worksheet not found – cannot validate password.
                    return false;
                }

                // Verify the supplied password against the worksheet protection.
                return worksheet.Protection.VerifyPassword(password);
            }
            catch (Exception)
            {
                // Return false on any exception for safety.
                return false;
            }
        }
    }

    // Entry point for console execution.
    public class Program
    {
        public static void Main(string[] args)
        {
            // Example usage; replace with actual values or command‑line arguments.
            string filePath = "sample.xlsx";
            string sheetName = "Sheet1";
            string password = "myPassword";

            bool isValid = WorksheetPasswordValidator.ValidatePassword(filePath, sheetName, password);
            Console.WriteLine(isValid
                ? "Password is valid for the worksheet."
                : "Password is invalid or worksheet/file not found.");
        }
    }
}
