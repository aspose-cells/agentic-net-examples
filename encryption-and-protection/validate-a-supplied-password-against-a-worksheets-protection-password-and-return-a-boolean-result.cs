// Title: Validate an Excel worksheet's protection password with Aspose.Cells for .NET (C#)
// Common Searches: C# code to check if a worksheet password is correct using Aspose.Cells | How to programmatically verify Excel sheet protection password in .NET | Aspose.Cells unprotect worksheet without exception and confirm password | Validate worksheet protection password from file path and sheet name in C# | Determine if an Excel worksheet is protected with a specific password using Aspose.Cells API
// Tags: Aspose.Cells worksheet password validation | C# unprotect Excel worksheet programmatically | Excel sheet protection check using Aspose | load workbook and verify worksheet password | validate worksheet protection .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    // Loads a workbook from a given file, retrieves the specified worksheet, and attempts to unprotect it with the supplied password. Returns true when the password matches the worksheet's protection password and false otherwise, while handling missing files and invalid worksheet names.
    public class WorksheetProtectionValidator
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="sheetName">Name of the worksheet to validate.</param>
        /// <param name="password">Password to test.</param>
        /// <returns>True if the password matches the worksheet's protection password; otherwise false.</returns>
        public static bool ValidateWorksheetPassword(string filePath, string sheetName, string password)
        {
            // Ensure the file exists before attempting to load it.
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Retrieve the worksheet by name.
            Worksheet worksheet = workbook.Worksheets[sheetName];
            if (worksheet == null)
                throw new ArgumentException($"Worksheet '{sheetName}' does not exist in the workbook.");

            // Attempt to unprotect using the supplied password.
            // If the password is correct, no exception is thrown.
            try
            {
                worksheet.Unprotect(password);
                return true;
            }
            catch
            {
                // Wrong password.
                return false;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – adjust these values as needed.
            string filePath = "sample.xlsx";
            string sheetName = "Sheet1";
            string password = "test";

            try
            {
                bool isValid = WorksheetProtectionValidator.ValidateWorksheetPassword(filePath, sheetName, password);
                Console.WriteLine($"Password validation result: {isValid}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
