// Title: How to detect password protection on an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to verify if an .xlsx file is password protected with Aspose.Cells | Aspose.Cells load workbook without password to test encryption status | Determine if Excel file requires a password before opening in .NET | Check Excel file encryption using LoadOptions and exception handling in C# | Identify password‑protected Excel workbook programmatically with Aspose.Cells
// Tags: aspocells detect excel password protection | c# check encrypted xlsx workbook | loadoptions auto without password aspocells | exception handling protected excel file aspocells | excel workbook encryption detection .net

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The IsPasswordProtected method confirms the file exists, then attempts to load the workbook with LoadOptions set to Auto and no password. If loading succeeds, the workbook is not password‑protected. If Aspose.Cells throws an exception containing keywords such as "password" or "encrypted", the method returns true, indicating the file is protected. Other exceptions are treated as non‑protected and result in false.
    public static class ExcelPasswordChecker
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>True if the file is password protected; otherwise false.</returns>
        public static bool IsPasswordProtected(string filePath)
        {
            // Ensure the file exists before attempting to load it.
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            try
            {
                // Attempt to load the workbook without providing a password.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                Workbook workbook = new Workbook(filePath, loadOptions);

                // If loading succeeds, the file is not password protected (or uses an empty password).
                return false;
            }
            catch (Exception ex)
            {
                // Aspose.Cells throws an exception when a password is required.
                // Check the message for typical password‑related keywords.
                if (ex.Message != null &&
                    (ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     ex.Message.IndexOf("encrypted", StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return true;
                }

                // For other exceptions, assume the file is not password protected.
                return false;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide the Excel file path as a command‑line argument or use a default.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            try
            {
                bool isProtected = ExcelPasswordChecker.IsPasswordProtected(filePath);
                Console.WriteLine($"File \"{filePath}\" password protected: {isProtected}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
