// Title: Determine whether an Excel workbook is password‑protected (encrypted) with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if Excel file is encrypted without a password | How to programmatically detect password protection on .xlsx using Aspose.Cells | C# method to verify workbook encryption status with Aspose.Cells API | Identify encrypted Excel workbook before opening it with Aspose.Cells | Detect if an Excel workbook requires a password using Aspose.Cells in .NET
// Tags: encrypted workbook detection Aspose.Cells | password‑protected Excel check .NET | CellsException handling for encrypted files | load workbook without password Aspose.Cells | C# workbook encryption status verification

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookUtility
{
    // The IsWorkbookEncrypted function validates the file path, attempts to load the workbook with Aspose.Cells without supplying a password, and returns true when a CellsException is thrown (indicating the workbook is encrypted); otherwise it returns false, re‑throwing any unrelated exceptions.
    public static class WorkbookHelper
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>True if the workbook is encrypted; otherwise, false.</returns>
        public static bool IsWorkbookEncrypted(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Workbook file not found: {filePath}");

            try
            {
                // Attempt to load without a password. If the file is encrypted, a CellsException is thrown.
                var wb = new Workbook(filePath);
                return false; // Loaded successfully, not encrypted.
            }
            catch (CellsException)
            {
                // Any CellsException during load without a password indicates encryption.
                return true;
            }
            catch (Exception)
            {
                // Re‑throw any other unexpected exceptions.
                throw;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide a path as the first argument.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the workbook file path as an argument.");
                return;
            }

            string path = args[0];

            try
            {
                bool encrypted = WorkbookHelper.IsWorkbookEncrypted(path);
                Console.WriteLine($"Workbook encrypted: {encrypted}");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
