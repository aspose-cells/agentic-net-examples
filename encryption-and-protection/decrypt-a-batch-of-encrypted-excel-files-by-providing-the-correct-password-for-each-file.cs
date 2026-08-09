// Title: C# Batch Decrypt Password‑Protected Excel Files with Aspose.Cells
// Description: A C# utility that iterates a dictionary of encrypted Excel file paths and passwords, detects each workbook’s format, opens it with LoadOptions, clears the password, and saves an unencrypted copy (preserving the original extension) to a target folder. Supports XLSX, XLS, CSV, PDF, ODS and can be extended for additional formats.
// Keywords: Aspose.Cells | C# batch decrypt Excel | remove Excel password .NET | detect Excel file format | save unencrypted workbook | Excel encryption removal | bulk Excel decryption | LoadOptions password | SaveFormat mapping | Aspose.Cells example GitHub
// Common Searches: how to remove password from multiple Excel files using Aspose.Cells | C# code to batch decrypt .xlsx and .xls files | Aspose.Cells bulk decryption tutorial | detect Excel file type before saving Aspose.Cells | save decrypted workbook with original extension C# | Aspose.Cells password removal for CSV files | batch Excel decryption script .NET
// Developer Intent: Decrypt many password‑protected Excel workbooks and write unprotected copies to a chosen directory.
// Use Cases: Automate nightly processing of archived encrypted reports before analytics. | Prepare a set of client‑provided protected spreadsheets for data migration. | Integrate into a document‑management workflow that requires password‑free files for OCR or conversion. | Extend to convert decrypted files to other formats (PDF, CSV) in bulk.
// AI Prompts: Write C# code using Aspose.Cells to bulk decrypt Excel files given a file‑path‑to‑password map and save them with a '_decrypted' suffix. | Show how to add support for Xlsb and Ooxml formats in the GetSaveFormat method while keeping original extensions. | Create robust error handling for missing files, wrong passwords, and unsupported formats in the batch decryption routine. | Generate a PowerShell wrapper that calls the C# batch decryptor for scheduled tasks. | Explain the performance impact of loading many encrypted workbooks and suggest optimizations.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A C# utility that iterates a dictionary of encrypted Excel file paths and passwords, detects each workbook’s format, opens it with LoadOptions, clears the password, and saves an unencrypted copy (preserving the original extension) to a target folder. Supports XLSX, XLS, CSV, PDF, ODS and can be extended for additional formats.
public class ExcelBatchDecryptor
{
    // Decrypts a batch of encrypted Excel files.
    // filePasswordMap: key = full path of encrypted file, value = password for that file.
    // outputFolder: folder where decrypted files will be saved.
    public void DecryptFiles(Dictionary<string, string> filePasswordMap, string outputFolder)
    {
        // Ensure output directory exists
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        foreach (var kvp in filePasswordMap)
        {
            string encryptedPath = kvp.Key;
            string password = kvp.Value;

            try
            {
                // Verify source file exists
                if (!File.Exists(encryptedPath))
                {
                    Console.WriteLine($"Source file not found: {encryptedPath}");
                    continue;
                }

                // Detect the original file format (needed for correct SaveFormat)
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(encryptedPath);
                SaveFormat saveFormat = GetSaveFormat(formatInfo.FileFormatType);

                // Load the encrypted workbook using the password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = password
                };
                Workbook workbook = new Workbook(encryptedPath, loadOptions);

                // Remove encryption by clearing the password property
                workbook.Settings.Password = null;

                // Build output file name (same name with "_decrypted" suffix)
                string fileName = Path.GetFileNameWithoutExtension(encryptedPath);
                string extension = Path.GetExtension(encryptedPath);
                string decryptedPath = Path.Combine(outputFolder, $"{fileName}_decrypted{extension}");

                // Save the workbook without password
                workbook.Save(decryptedPath, saveFormat);
                Console.WriteLine($"Decrypted file saved: {decryptedPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{encryptedPath}': {ex.Message}");
            }
        }
    }

    // Maps FileFormatType to corresponding SaveFormat enum value.
    private SaveFormat GetSaveFormat(FileFormatType fileFormatType)
    {
        // Most common mappings; extend as needed.
        switch (fileFormatType)
        {
            case FileFormatType.Xlsx:
                return SaveFormat.Xlsx;
            case FileFormatType.Csv:
                return SaveFormat.Csv;
            case FileFormatType.Pdf:
                return SaveFormat.Pdf;
            case FileFormatType.Ods:
                return SaveFormat.Ods;
            // For older Excel formats (e.g., .xls) fall back to Excel97To2003
            default:
                return SaveFormat.Excel97To2003;
        }
    }
}

// Example usage
public class Program
{
    public static void Main()
    {
        try
        {
            // Map of encrypted files and their passwords
            var files = new Dictionary<string, string>
            {
                { @"C:\Encrypted\Report1.xlsx", "Pass123" },
                { @"C:\Encrypted\Report2.xls", "Secret!" },
                { @"C:\Encrypted\Data.csv", "CsvPwd" }
            };

            string outputFolder = @"C:\Decrypted";

            var decryptor = new ExcelBatchDecryptor();
            decryptor.DecryptFiles(files, outputFolder);

            Console.WriteLine("Decryption completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
