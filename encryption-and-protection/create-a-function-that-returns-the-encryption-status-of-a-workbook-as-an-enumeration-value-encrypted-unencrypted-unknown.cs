// Title: Determine Excel workbook encryption status (Encrypted, Unencrypted, Unknown) with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells check if an Excel file is password protected without opening it | How to return an enumeration for workbook encryption state using Aspose.Cells | Detect encrypted Excel workbook in .NET using LoadOptions and exception handling | Get encryption status of .xlsx file with Aspose.Cells API | Identify unknown workbook protection errors in C# Aspose.Cells
// Tags: Aspose.Cells workbook encryption detection | C# Excel encryption status enumeration | LoadOptions password handling Aspose.Cells | CellsException password required detection | Determine encrypted vs unencrypted Excel file

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionChecker
{
    // Enumeration representing the encryption status of a workbook
    public enum EncryptionStatus
    {
        Encrypted,
        Unencrypted,
        Unknown
    }

    // The GetEncryptionStatus method validates the file path, attempts to load the workbook with default LoadOptions, returns Unencrypted on successful load, catches a CellsException containing "password" to indicate Encrypted, and returns Unknown for any other exception or unsupported format.
    public static class EncryptionHelper
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>EncryptionStatus value indicating Encrypted, Unencrypted, or Unknown.</returns>
        public static EncryptionStatus GetEncryptionStatus(string filePath)
        {
            // Ensure the file exists before attempting to load it.
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");

            try
            {
                // Attempt to load the workbook without a password.
                var loadOptions = new LoadOptions(); // default options, no password
                var workbook = new Workbook(filePath, loadOptions);

                // Loading succeeded, so the workbook is not encrypted.
                return EncryptionStatus.Unencrypted;
            }
            catch (CellsException ex)
            {
                // Aspose.Cells throws a CellsException when a password is required.
                // Detect this via the exception message.
                if (ex.Message != null &&
                    ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return EncryptionStatus.Encrypted;
                }

                // Any other CellsException is treated as unknown status.
                return EncryptionStatus.Unknown;
            }
            catch (Exception)
            {
                // Non‑Aspose.Cells exceptions (e.g., invalid format) result in Unknown.
                return EncryptionStatus.Unknown;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string filePath;

                if (args.Length > 0)
                {
                    filePath = args[0];
                }
                else
                {
                    Console.Write("Enter workbook file path: ");
                    filePath = Console.ReadLine();
                }

                var status = EncryptionHelper.GetEncryptionStatus(filePath);
                Console.WriteLine($"Encryption status: {status}");
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
