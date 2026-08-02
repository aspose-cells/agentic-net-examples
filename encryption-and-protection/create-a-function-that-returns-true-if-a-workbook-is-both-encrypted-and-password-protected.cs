// Title: C# – Check if an Excel workbook is encrypted and password‑protected with Aspose.Cells
// Description: A C# helper method that validates the file path, detects encryption via FileFormatUtil.DetectFileFormat, loads the workbook with LoadOptions when required, and uses Workbook.IsWorkbookProtectedWithPassword to return true only when the workbook is both encrypted and password‑protected.
// Keywords: Aspose.Cells | C# workbook encryption detection | FileFormatUtil DetectFileFormat | Workbook.IsWorkbookProtectedWithPassword | encrypted Excel file check | password protected workbook .NET | Excel security validation Aspose | load encrypted workbook Aspose.Cells
// Common Searches: Aspose.Cells detect encrypted Excel file | C# check if workbook is password protected | IsWorkbookProtectedWithPassword example | determine if Excel file is encrypted and protected using Aspose.Cells | detect encryption before loading workbook Aspose.Cells .NET
// Developer Intent: The developer wants to know whether a given workbook file is both encrypted and protected with a password.
// Use Cases: Validate security of uploaded Excel files before processing them in a web service. | Reject or flag workbooks that lack encryption or password protection in compliance‑driven applications. | Log encryption and protection status for a batch of workbooks during nightly audits.
// AI Prompts: Write unit tests for IsWorkbookEncryptedAndProtected covering encrypted‑protected, encrypted‑unprotected, unencrypted‑protected, and missing file scenarios. | Generate C# code that scans a directory, calls IsWorkbookEncryptedAndProtected for each file, and writes the results to a CSV log. | Explain the difference between Workbook.IsWorkbookProtectedWithPassword and individual worksheet protection settings in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // A C# helper method that validates the file path, detects encryption via FileFormatUtil.DetectFileFormat, loads the workbook with LoadOptions when required, and uses Workbook.IsWorkbookProtectedWithPassword to return true only when the workbook is both encrypted and password‑protected.
    public static class WorkbookProtectionHelper
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <param name="password">
        /// Password to open the workbook if it is encrypted.
        /// Pass null or empty string if the workbook is not encrypted.
        /// </param>
        /// <returns>True when both conditions are met; otherwise false.</returns>
        public static bool IsWorkbookEncryptedAndProtected(string filePath, string password = null)
        {
            // Ensure the file exists before attempting any operation.
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                // Detect whether the file is encrypted without loading the workbook.
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                bool isEncrypted = formatInfo.IsEncrypted;

                // Load the workbook. Supply the password if it is encrypted.
                Workbook workbook;
                if (isEncrypted)
                {
                    var loadOptions = new LoadOptions { Password = password };
                    workbook = new Workbook(filePath, loadOptions);
                }
                else
                {
                    workbook = new Workbook(filePath);
                }

                // Check if the workbook structure or window is protected with a password.
                bool isProtected = workbook.IsWorkbookProtectedWithPassword;

                // Return true only when both conditions are satisfied.
                return isEncrypted && isProtected;
            }
            catch (Exception ex)
            {
                // Log the exception and return false to indicate failure.
                Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
                return false;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // args[0] = path to workbook, args[1] = optional password
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";
            string password = args.Length > 1 ? args[1] : null;

            bool result = WorkbookProtectionHelper.IsWorkbookEncryptedAndProtected(filePath, password);
            Console.WriteLine($"Workbook encrypted and protected: {result}");
        }
    }
}
