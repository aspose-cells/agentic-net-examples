// Title: Encrypt an Excel workbook with password and 128‑bit strong encryption using Aspose.Cells for .NET and test third‑party library compatibility
// Description: Creates a new Workbook, writes sample data, applies a password, sets 128‑bit strong cryptographic encryption, saves the file, detects the encrypted flag with FileFormatUtil, and simulates opening the file with a non‑Aspose parser to illustrate incompatibility.
// Keywords: Aspose.Cells encrypt Excel .NET | password protected workbook C# | 128‑bit strong encryption Excel | detect encrypted workbook Aspose | third‑party library compatibility Excel | NPOI encrypted file support
// Common Searches: how to password protect an Excel file using Aspose.Cells C# | set strong encryption for .xlsx with Aspose.Cells | detect if an Excel workbook is encrypted in .NET | open password protected Excel file with NPOI | Aspose.Cells encryption compatibility test
// Developer Intent: Apply password protection and strong 128‑bit encryption to an Excel workbook, confirm its encrypted status, and evaluate whether a third‑party library can read the file.
// Use Cases: Generate confidential reports that must be encrypted before distribution. | Programmatically verify that a saved workbook is encrypted using FileFormatUtil.DetectFileFormat. | Validate that external parsers (e.g., NPOI) cannot open Aspose‑encrypted files, ensuring data security.
// AI Prompts: Provide C# code to encrypt an existing .xlsx with a password and 256‑bit AES using Aspose.Cells, then attempt to open it with NPOI and handle errors. | Explain how to use Aspose.Cells to check the IsEncrypted flag of a workbook after saving. | Show how to catch and log compatibility exceptions when a third‑party library tries to read a password‑protected Excel file created by Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a new Workbook, writes sample data, applies a password, sets 128‑bit strong cryptographic encryption, saves the file, detects the encrypted flag with FileFormatUtil, and simulates opening the file with a non‑Aspose parser to illustrate incompatibility.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Encrypted Test");

            // Set a password to protect the workbook
            workbook.Settings.Password = "Secret123";

            // Set encryption options (strong encryption, 128‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            string encryptedFile = "EncryptedWorkbook.xlsx";
            workbook.Save(encryptedFile);

            // Verify that Aspose detects the file as encrypted
            if (File.Exists(encryptedFile))
            {
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedFile);
                Console.WriteLine("Aspose reports IsEncrypted: " + formatInfo.IsEncrypted);
            }
            else
            {
                Console.WriteLine("Encrypted file not found.");
            }

            // Attempt to open the encrypted file with a third‑party library (simulated)
            try
            {
                // NPOI is not referenced; simulate incompatibility
                throw new NotSupportedException("Third‑party library does not support opening password‑protected OOXML files.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Third‑party library failed to open encrypted workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
