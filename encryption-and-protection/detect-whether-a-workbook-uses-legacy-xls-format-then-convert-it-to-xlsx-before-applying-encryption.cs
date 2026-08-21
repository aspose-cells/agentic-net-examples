// Title: Convert legacy XLS to XLSX and encrypt with password using Aspose.Cells for .NET
// Description: Detect a legacy .xls workbook, convert it to .xlsx via a temporary file, then apply a password and 128‑bit strong encryption before saving the protected file.
// Keywords: Aspose.Cells | C# | .NET | convert XLS to XLSX | encrypt Excel workbook | password protection | strong encryption | AES 128 | temporary file cleanup | legacy Excel format detection
// Common Searches: How to convert .xls to .xlsx and password protect with Aspose.Cells | Aspose.Cells encrypt workbook with strong cryptographic provider | Detect legacy Excel file before applying encryption in C# | Save encrypted XLSX to custom folder using Aspose.Cells | Batch convert and encrypt mixed .xls/.xlsx files .NET
// Developer Intent: Identify whether an input workbook is an old .xls file, transform it to .xlsx if needed, and save the result with password‑based strong encryption.
// Use Cases: Upgrade archived Excel reports to encrypted .xlsx files for secure sharing. | Automate processing of mixed‑format spreadsheets, ensuring every output is password‑protected. | Integrate conversion‑encryption logic into a file‑upload service that only accepts encrypted .xlsx files.
// AI Prompts: Write C# code with Aspose.Cells that checks a file extension, converts .xls to .xlsx, applies a password and 128‑bit AES encryption, and removes temporary files. | Show an Aspose.Cells example that encrypts an existing .xlsx workbook, creates missing output directories, and uses the StrongCryptographicProvider. | Explain how to extend the sample to assign different passwords to individual worksheets while keeping the whole workbook encrypted.

using System;
using System.IO;
using Aspose.Cells;

// Detect a legacy .xls workbook, convert it to .xlsx via a temporary file, then apply a password and 128‑bit strong encryption before saving the protected file.
public class LegacyXlsToXlsxEncryptor
{
    /// <param name="sourcePath">Path to the original workbook (XLS or XLSX).</param>
    /// <param name="encryptedPath">Path where the encrypted XLSX workbook will be saved.</param>
    /// <param name="password">Password to protect the workbook.</param>
    public static void ConvertAndEncrypt(string sourcePath, string encryptedPath, string password)
    {
        // Verify that the source file exists to avoid FileNotFoundException.
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException($"Source file not found: {sourcePath}");
        }

        // Determine whether the source is a legacy XLS file based on its extension.
        bool isLegacyXls = sourcePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase);

        Workbook workbook = null;
        string tempXlsxPath = null;

        try
        {
            // Load the workbook (Aspose.Cells automatically handles the format).
            workbook = new Workbook(sourcePath);

            // If the source was XLS, convert it to XLSX first.
            if (isLegacyXls)
            {
                // Save the workbook as XLSX to a temporary location.
                tempXlsxPath = Path.Combine(Path.GetTempPath(),
                                            Guid.NewGuid().ToString("N") + ".xlsx");
                workbook.Save(tempXlsxPath, SaveFormat.Xlsx);

                // Reload the newly saved XLSX for further processing.
                workbook.Dispose();
                workbook = new Workbook(tempXlsxPath);
            }

            // Apply password protection.
            workbook.Settings.Password = password;

            // Set stronger encryption options (relevant for XLSX).
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(encryptedPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the encrypted workbook as XLSX.
            workbook.Save(encryptedPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            // Log or rethrow as needed.
            Console.Error.WriteLine($"Error during conversion/encryption: {ex.Message}");
            throw;
        }
        finally
        {
            // Clean up resources.
            workbook?.Dispose();

            // Delete temporary file if it was created.
            if (tempXlsxPath != null && File.Exists(tempXlsxPath))
            {
                try { File.Delete(tempXlsxPath); } catch { /* ignore cleanup errors */ }
            }
        }
    }

    // Example usage.
    public static void Main()
    {
        try
        {
            string sourceFile = "LegacyWorkbook.xls";          // Input file (could be .xls or .xlsx)
            string encryptedFile = "EncryptedWorkbook.xlsx";   // Desired output file
            string password = "MySecurePassword";

            ConvertAndEncrypt(sourceFile, encryptedFile, password);

            Console.WriteLine("Conversion and encryption completed successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Operation failed: {ex.Message}");
        }
    }
}
