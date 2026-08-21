// Title: C# Utility to Extract Excel Encryption Metadata with Aspose.Cells
// Description: A console‑based C# example that uses Aspose.Cells' FileFormatUtil to detect an Excel file's format, report encryption flags (IsEncrypted, IsProtectedByRMS), and optionally load the workbook with LoadOptions while handling missing or incorrect passwords. Ideal for quick audits of password‑protected or RMS‑protected spreadsheets.
// Keywords: Aspose.Cells encryption metadata | C# detect Excel password protection | FileFormatUtil IsEncrypted | IsProtectedByRMS Aspose | LoadOptions encrypted workbook | Excel file security audit .NET | console utility Aspose.Cells | read Excel file protection status
// Common Searches: how to check if an Excel file is encrypted using Aspose.Cells | C# code to read encryption flags from XLSX | display IsEncrypted and IsProtectedByRMS properties | load encrypted workbook with password Aspose.Cells | detect RMS protection in Excel with .NET
// Developer Intent: Identify and display encryption‑related properties of an Excel workbook and optionally attempt to open it with a password.
// Use Cases: Validate incoming Excel files for password protection before processing. | Log encryption and RMS status for compliance or security audits. | Programmatically attempt to open a protected workbook when the password is known. | Integrate a quick‑look utility into CI pipelines to flag encrypted spreadsheets.
// AI Prompts: Write C# code that uses Aspose.Cells to detect encryption and RMS protection in an Excel file. | Show how to load an encrypted workbook with a password using LoadOptions in Aspose.Cells. | Explain error handling for missing or incorrect passwords when opening encrypted Excel files with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsEncryptionMetadataUtility
{
    // Utility class to extract and display encryption related metadata from an Excel file
    // A console‑based C# example that uses Aspose.Cells' FileFormatUtil to detect an Excel file's format, report encryption flags (IsEncrypted, IsProtectedByRMS), and optionally load the workbook with LoadOptions while handling missing or incorrect passwords. Ideal for quick audits of password‑protected or RMS‑protected spreadsheets.
    public static class EncryptionMetadataUtility
    {
        // Displays encryption metadata for the specified Excel file
        public static void DisplayEncryptionMetadata(string filePath)
        {
            // Detect file format and retrieve metadata information
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

            // Output basic format information
            Console.WriteLine("=== Encryption Metadata ===");
            Console.WriteLine($"File Path               : {filePath}");
            Console.WriteLine($"Detected Load Format    : {formatInfo.LoadFormat}");
            Console.WriteLine($"Detected File Format    : {formatInfo.FileFormatType}");
            Console.WriteLine($"Is Encrypted            : {formatInfo.IsEncrypted}");
            Console.WriteLine($"Is Protected By RMS     : {formatInfo.IsProtectedByRMS}");
            Console.WriteLine();

            // If the file is encrypted, attempt to load it with a password (if known)
            if (formatInfo.IsEncrypted)
            {
                // Example: you can set a password here if you have it.
                // For demonstration, we will try without a password to show the exception handling.
                try
                {
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                    // Uncomment and set the password if you have it:
                    // loadOptions.Password = "your_password";

                    Workbook workbook = new Workbook(filePath, loadOptions);
                    // After successful load, we can check the workbook settings
                    Console.WriteLine("Workbook loaded successfully.");
                    Console.WriteLine($"Workbook Settings.IsEncrypted : {workbook.Settings.IsEncrypted}");
                }
                catch (Exception ex)
                {
                    // Expected when password is missing or incorrect
                    Console.WriteLine("Unable to load the encrypted workbook without a valid password.");
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("The file is not encrypted.");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            // Replace with the path to your Excel file
            string excelFilePath = "sample.xlsx";

            EncryptionMetadataUtility.DisplayEncryptionMetadata(excelFilePath);
        }
    }
}
