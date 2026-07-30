// Title: Batch Decrypt, Sanitize, and Re‑Encrypt Excel Files with Aspose.Cells for .NET
// Description: A C# utility that scans a folder for .xlsx workbooks, detects encryption, opens each file with the old password, removes personal data (e.g., comments, SSN placeholders), applies a new password, and saves the cleaned files to an output directory for GDPR, HIPAA, or other compliance needs.
// Keywords: Aspose.Cells | C# batch decrypt Excel | Excel data sanitization | remove personal information from workbook | re‑encrypt Excel with new password | FileFormatUtil | LoadOptions password | compliance automation | SSN redaction | GDPR Excel protection
// Common Searches: how to batch decrypt encrypted Excel files using Aspose.Cells | C# code to remove personal information from multiple workbooks | re‑encrypt Excel files with a new password programmatically | Aspose.Cells example for sanitizing and protecting spreadsheets | automate GDPR redaction of Excel files in .NET
// Developer Intent: Decrypt each workbook, redact sensitive content, and save it encrypted with a new password.
// Use Cases: Regulatory compliance: scan encrypted financial reports, redact SSNs, and protect them with a corporate password. | Secure external sharing: strip author metadata and comments before distributing spreadsheets to partners. | CI/CD pipeline: automatically sanitize and re‑secure generated dashboards each night before release.
// AI Prompts: Write C# code using Aspose.Cells to open encrypted .xlsx files with a given password, replace all occurrences of a sensitive keyword, and save the workbook with a new password. | Show how to handle both encrypted and plain Excel files in a batch loop using FileFormatUtil and LoadOptions. | Suggest extensions for the sanitization step, such as clearing specific worksheet ranges or deleting comments that contain personal data.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A C# utility that scans a folder for .xlsx workbooks, detects encryption, opens each file with the old password, removes personal data (e.g., comments, SSN placeholders), applies a new password, and saves the cleaned files to an output directory for GDPR, HIPAA, or other compliance needs.
class BatchDecryptSanitizeReencrypt
{
    static void Main()
    {
        // Folder containing the source Excel files
        string inputFolder = @"C:\InputExcel";
        // Folder where the processed files will be saved
        string outputFolder = @"C:\OutputExcel";

        // Old password used to open encrypted files (if any)
        string oldPassword = "OldPassword123";
        // New password to protect the processed files
        string newPassword = "NewSecurePassword456";

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Verify input directory exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Process each Excel file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found (skipped): {filePath}");
                    continue;
                }

                // Detect file format and encryption status
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

                // Load workbook with appropriate options
                Workbook workbook;
                if (formatInfo.IsEncrypted)
                {
                    // Load encrypted workbook using the old password
                    LoadOptions loadOptions = new LoadOptions
                    {
                        Password = oldPassword
                    };
                    workbook = new Workbook(filePath, loadOptions);
                }
                else
                {
                    // Load unencrypted workbook
                    workbook = new Workbook(filePath);
                }

                // ---------- Data Sanitization ----------
                // Remove personal information such as author names in comments
                workbook.RemovePersonalInformation();

                // Example: replace any occurrence of a placeholder "SSN" with "REDACTED"
                workbook.Replace("SSN", "REDACTED");

                // Additional sanitization logic can be added here (e.g., clearing specific ranges)

                // ---------- Re‑encryption ----------
                // Set the new password for the workbook
                workbook.Settings.Password = newPassword;

                // Save the sanitized and re‑encrypted workbook to the output folder
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                workbook.Save(outputPath);

                // Clean up
                workbook.Dispose();

                Console.WriteLine($"Processed and saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
