// Title: Check if a worksheet’s protection password matches a given string using Aspose.Cells for .NET (C#)
// Common Searches: C# code to verify an Excel worksheet protection password with Aspose.Cells | How to compare a supplied password to a worksheet’s protection in Aspose.Cells .NET | Validate worksheet password before editing using Aspose.Cells library | Aspose.Cells example for verifying worksheet protection password | Programmatically check worksheet protection password in C#
// Tags: worksheet protection password verification Aspose.Cells .NET | Aspose.Cells verify worksheet password C# | load workbook and check sheet protection password | Excel worksheet password validation using Aspose.Cells API | C# Aspose.Cells worksheet protection check

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetPasswordValidatorApp
{
    // The ValidateWorksheetPassword method loads a workbook from a file, locates the specified worksheet, and uses Worksheet.Protection.VerifyPassword to determine whether the provided password matches the sheet's protection password, returning true on success and false for missing files, sheets, or mismatches.
    public class WorksheetPasswordValidator
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="sheetName">Name of the worksheet to check.</param>
        /// <param name="password">Password to validate.</param>
        /// <returns>True if the password matches the worksheet's protection password; otherwise false.</returns>
        public bool ValidateWorksheetPassword(string filePath, string sheetName, string password)
        {
            try
            {
                // Ensure the file exists before attempting to load.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return false;
                }

                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(filePath);

                // Retrieve the worksheet by name.
                Worksheet worksheet = workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    Console.WriteLine($"Worksheet '{sheetName}' not found.");
                    return false;
                }

                // Verify the supplied password against the worksheet's protection password.
                return worksheet.Protection.VerifyPassword(password);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"Error during validation: {ex.Message}");
                return false;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – replace with actual values as needed.
            string filePath = "sample.xlsx";
            string sheetName = "Sheet1";
            string password = "yourPassword";

            var validator = new WorksheetPasswordValidator();
            bool isValid = validator.ValidateWorksheetPassword(filePath, sheetName, password);

            Console.WriteLine(isValid ? "Password is correct." : "Password is incorrect or validation failed.");
        }
    }
}
