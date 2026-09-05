// Title: Determine Excel workbook encryption status (Encrypted, Unencrypted, Unknown) with Aspose.Cells in C#
// AI Prompts: Implement a C# method that uses Aspose.Cells LoadOptions to open an Excel file and returns an EncryptionStatus enum based on whether a password‑required exception occurs. | Create NUnit tests for EncryptionHelper.GetEncryptionStatus covering encrypted, unencrypted, missing, and corrupted workbook scenarios. | Add logging to GetEncryptionStatus to capture the cause of an Unknown result while preserving the existing enum return values.
// Common Searches: Aspose.Cells C# check if an .xlsx file is password protected without providing a password | How to programmatically detect encrypted Excel workbook using Aspose.Cells LoadOptions | C# get encryption state of Excel file (encrypted, unencrypted, unknown) with Aspose.Cells | Determine if an Excel workbook requires a password before opening in .NET
// Tags: Aspose.Cells detect workbook encryption | C# load Excel file without password | EncryptionStatus enum Aspose.Cells | LoadOptions Auto format password exception handling | CellsException password error detection

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

    // Provides an EncryptionStatus enum and a static EncryptionHelper.GetEncryptionStatus method that attempts to load an Excel workbook with Aspose.Cells. A successful load returns Unencrypted, a password‑related CellsException returns Encrypted, and any other failure returns Unknown.
    public static class EncryptionHelper
    {
        /// <param name="filePath">Full path to the workbook file.</param>
        /// <returns>EncryptionStatus value indicating Encrypted, Unencrypted, or Unknown.</returns>
        public static EncryptionStatus GetEncryptionStatus(string filePath)
        {
            // Verify that the file exists before attempting to load it
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return EncryptionStatus.Unknown;

            try
            {
                // Load the workbook without providing a password.
                // If the file is encrypted and a password is required, an exception will be thrown.
                var loadOptions = new LoadOptions(LoadFormat.Auto);
                var workbook = new Workbook(filePath, loadOptions);

                // If loading succeeded, the workbook is not password‑protected.
                // Aspose.Cells versions prior to 23.9 do not expose an IsEncrypted property,
                // so we treat a successful load as unencrypted.
                return EncryptionStatus.Unencrypted;
            }
            catch (CellsException ex)
            {
                // Detect password‑required errors via the exception message.
                if (!string.IsNullOrEmpty(ex.Message) &&
                    ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return EncryptionStatus.Encrypted;
                }

                // Any other CellsException means the status cannot be determined reliably.
                return EncryptionStatus.Unknown;
            }
            catch (Exception)
            {
                // Non‑Aspose.Cells exceptions (e.g., I/O errors) result in an Unknown status.
                return EncryptionStatus.Unknown;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide the workbook path as a command‑line argument or use a default.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            EncryptionStatus status = EncryptionHelper.GetEncryptionStatus(filePath);
            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"Encryption status: {status}");
        }
    }
}
