// Title: Determine whether an Excel .xlsx workbook is protected by a modify password using Aspose.Cells in C#
// Common Searches: how to detect a modify password on an Excel file using Aspose.Cells in C# | C# Aspose.Cells check if workbook requires a modify password before editing | programmatically verify Excel .xlsx modify protection with Aspose.Cells | using Aspose.Cells Unprotect method to test workbook modify password validity | determine if Excel workbook is locked for modifications in .NET
// Tags: Aspose.Cells workbook unprotect operation | C# Excel modify password verification | Aspose.Cells workbook protection status query | .xlsx modify password check using .NET | Excel file modification lock detection

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookPasswordCheckerApp
{
    // IsModifyPasswordRequired loads the workbook with Aspose.Cells, attempts to unprotect it using the supplied password, and returns true when a modify password exists or the password is incorrect; otherwise it returns false. It throws FileNotFoundException for missing files and wraps loading errors in InvalidOperationException.
    public class WorkbookPasswordChecker
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <param name="password">Password to test.</param>
        /// <returns>True if a modify password is required (or the supplied password is incorrect); otherwise false.</returns>
        public static bool IsModifyPasswordRequired(string filePath, string password)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            Workbook workbook;
            try
            {
                // Load the workbook without a password (modify password does not affect opening)
                workbook = new Workbook(filePath);
            }
            catch (Exception ex)
            {
                // Wrap loading errors for clearer diagnostics
                throw new InvalidOperationException("Failed to load workbook.", ex);
            }

            try
            {
                // Attempt to unprotect using the supplied password.
                workbook.Unprotect(password);
                // No exception means either no modify password or the password is correct.
                return false;
            }
            catch
            {
                // Exception indicates a modify password exists and the supplied password is wrong.
                return true;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: file path and password can be passed via command line arguments.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";
            string password = args.Length > 1 ? args[1] : string.Empty;

            try
            {
                bool isRequired = WorkbookPasswordChecker.IsModifyPasswordRequired(filePath, password);
                Console.WriteLine(isRequired
                    ? "A modify password is required (or the supplied password is incorrect)."
                    : "No modify password required or the supplied password is correct.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"Error: {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
