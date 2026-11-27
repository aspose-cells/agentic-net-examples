// Title: C# function to retrieve Excel workbook encryption status (Encrypted, Unencrypted, Unknown) with Aspose.Cells
// Common Searches: Aspose.Cells how to check if an Excel workbook is password protected in C# | C# get encryption status of .xlsx file using Aspose.Cells without a password | Detect encrypted Excel workbook with Aspose.Cells LoadOptions Auto | Retrieve workbook encryption enum (Encrypted, Unencrypted, Unknown) in .NET | Handle CellsException for password‑protected Excel files in Aspose.Cells
// Tags: Aspose.Cells workbook encryption detection | C# GetEncryptionStatus method for Excel files | LoadOptions Auto passwordless workbook load Aspose.Cells | CellsException handling for password required Excel | Reflection IsEncrypted property Aspose.Cells | Excel encryption status enumeration in .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    public enum EncryptionStatus
    {
        Encrypted,
        Unencrypted,
        Unknown
    }

    // The GetEncryptionStatus method loads the workbook with LoadOptions.Auto, uses reflection to read the Workbook.IsEncrypted property when present, and catches CellsException messages containing "password" to return an EncryptionStatus enum (Encrypted, Unencrypted, Unknown), while gracefully handling missing files and other errors.
    public static class WorkbookHelper
    {
        public static EncryptionStatus GetEncryptionStatus(string filePath)
        {
            // Verify the file exists before attempting to load it.
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return EncryptionStatus.Unknown;

            try
            {
                // Load the workbook without providing a password.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Try to read the IsEncrypted property via reflection (available in newer versions).
                var isEncryptedProp = typeof(Workbook).GetProperty("IsEncrypted");
                if (isEncryptedProp != null && isEncryptedProp.PropertyType == typeof(bool))
                {
                    bool isEncrypted = (bool)isEncryptedProp.GetValue(workbook);
                    return isEncrypted ? EncryptionStatus.Encrypted : EncryptionStatus.Unencrypted;
                }

                // If the property is not present, successful load implies the file is not encrypted.
                return EncryptionStatus.Unencrypted;
            }
            catch (CellsException ex)
            {
                // Aspose.Cells throws an exception with a message containing "password" when a password is required.
                if (ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                    return EncryptionStatus.Encrypted;

                // Any other CellsException is treated as unknown.
                return EncryptionStatus.Unknown;
            }
            catch
            {
                // Non‑Cells exceptions (e.g., I/O errors) are treated as unknown.
                return EncryptionStatus.Unknown;
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage: provide the workbook path as the first argument.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            EncryptionStatus status = WorkbookHelper.GetEncryptionStatus(filePath);
            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"Encryption status: {status}");
        }
    }
}
