using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public static class WorkbookProtectionHelper
    {
        /// <summary>
        /// Determines whether the specified workbook is both encrypted (requires a password to open)
        /// and has its structure or window protected with a password.
        /// </summary>
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <param name="password">
        /// Password used to open the workbook. If the workbook is not encrypted this value is ignored.
        /// </param>
        /// <returns>True if the workbook is encrypted and also protected with a password; otherwise false.</returns>
        public static bool IsEncryptedAndPasswordProtected(string filePath, string password)
        {
            // Verify that the file exists before attempting to load it.
            if (!File.Exists(filePath))
                return false;

            try
            {
                // Load the workbook using the supplied password (if any).
                var loadOptions = new LoadOptions { Password = password };
                var workbook = new Workbook(filePath, loadOptions);

                // Workbook.Settings.IsEncrypted indicates whether a password is required to open the file.
                bool isEncrypted = workbook.Settings.IsEncrypted;

                // Workbook.IsWorkbookProtectedWithPassword indicates whether the workbook's structure
                // or window is protected with a password.
                bool isProtectedWithPassword = workbook.IsWorkbookProtectedWithPassword;

                // Return true only when both conditions are satisfied.
                return isEncrypted && isProtectedWithPassword;
            }
            catch (Exception)
            {
                // Any exception (e.g., wrong password, corrupted file) results in a false outcome.
                return false;
            }
        }
    }

    internal class Program
    {
        // Entry point required for a console application.
        private static void Main(string[] args)
        {
            // Example usage:
            // args[0] = path to workbook, args[1] = password (optional)
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the workbook file path as the first argument.");
                return;
            }

            string filePath = args[0];
            string password = args.Length > 1 ? args[1] : string.Empty;

            bool result = WorkbookProtectionHelper.IsEncryptedAndPasswordProtected(filePath, password);
            Console.WriteLine($"Workbook encrypted and password protected: {result}");
        }
    }
}