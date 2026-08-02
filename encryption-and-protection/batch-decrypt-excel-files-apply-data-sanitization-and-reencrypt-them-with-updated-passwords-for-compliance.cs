// Title: C# batch decrypt, sanitize, and re‑encrypt Excel files with Aspose.Cells
// Description: Scans a folder for .xlsx workbooks, detects encryption, opens each file with the old password, removes personal information, redacts specified strings, applies a new password, and saves the cleaned file to an output directory. Includes error handling for individual files and supports compliance‑driven workflows.
// Keywords: Aspose.Cells batch decrypt Excel | C# remove personal information Excel | re‑encrypt workbook new password | detect encrypted Excel file | sanitize Excel content programmatically | Excel compliance automation .NET
// Common Searches: batch decrypt and re‑encrypt Excel files C# | how to remove personal data from Excel using Aspose.Cells | change password of multiple encrypted workbooks programmatically | detect if an Excel workbook is password protected with Aspose.Cells | automate Excel sanitization and re‑encryption
// Developer Intent: Load encrypted Excel workbooks, cleanse sensitive data, and save them re‑protected with a new password in a single batch operation.
// Use Cases: Securely archive confidential reports by stripping author comments and redacting placeholders before applying a stronger organization‑wide password. | Ensure regulatory compliance for incoming Excel uploads by automatically removing personal identifiers and re‑encrypting the files. | Integrate into CI/CD pipelines to guarantee that all generated Excel artifacts are sanitized and protected before distribution.
// AI Prompts: Write C# code using Aspose.Cells to batch change passwords of encrypted Excel workbooks while preserving all worksheets and formulas. | Show how to extend the sanitization step to clear specific cell ranges based on a JSON configuration file. | Explain how to handle mixed Excel formats (xls, xlsx, xlsb) in the same batch process with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchDecryptSanitizeReencrypt
{
    // Scans a folder for .xlsx workbooks, detects encryption, opens each file with the old password, removes personal information, redacts specified strings, applies a new password, and saves the cleaned file to an output directory. Includes error handling for individual files and supports compliance‑driven workflows.
    class Program
    {
        static void Main(string[] args)
        {
            // Input parameters (could be replaced with args parsing)
            string inputFolder = @"C:\InputExcelFiles";
            string outputFolder = @"C:\SanitizedExcelFiles";
            string oldPassword = "oldPass123";
            string newPassword = "newSecurePass456";

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    // Detect file format and encryption status
                    FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                    bool isEncrypted = formatInfo.IsEncrypted;

                    // Load the workbook with appropriate load options
                    Workbook workbook;
                    if (isEncrypted)
                    {
                        // Use LoadOptions to supply the old password
                        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                        {
                            Password = oldPassword
                        };
                        workbook = new Workbook(filePath, loadOptions);
                    }
                    else
                    {
                        // Load without password
                        workbook = new Workbook(filePath);
                    }

                    // ------------------- Data Sanitization -------------------
                    // 1. Remove personal information such as author names in comments
                    workbook.RemovePersonalInformation();

                    // 2. Example string replacement to redact sensitive data
                    // Replace any occurrence of the placeholder "SensitiveData" with "REDACTED"
                    workbook.Replace("SensitiveData", "REDACTED");

                    // Additional sanitization logic can be added here (e.g., clearing specific ranges)

                    // ------------------- Re‑encrypt with new password -------------------
                    // Set the new password on the workbook settings
                    workbook.Settings.Password = newPassword;

                    // Save the sanitized and re‑encrypted workbook to the output folder
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                    workbook.Save(outputPath);

                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
