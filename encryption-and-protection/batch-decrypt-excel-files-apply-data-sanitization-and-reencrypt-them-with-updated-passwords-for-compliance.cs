// Title: C# Batch Decrypt, Sanitize, and Re‑encrypt Encrypted Excel Files with Aspose.Cells
// Description: A .NET console app that scans a folder of password‑protected .xlsx workbooks, opens each with the old password, removes personal data (e.g., comments, SSN values), applies a new password, and saves the cleaned files to an output directory. Includes robust error handling for compliance‑driven bulk processing.
// Keywords: Aspose.Cells | C# batch decrypt Excel | re‑encrypt workbook | Excel data sanitization | remove personal information Excel | redact SSN in Excel | password protected Excel files | GDPR compliance Excel | HIPAA Excel encryption | bulk Excel processing .NET
// Common Searches: How to open and decrypt password protected Excel files using Aspose.Cells C# | Batch replace sensitive text in Excel workbooks with Aspose.Cells | Re‑encrypt multiple Excel files with a new password in .NET | Remove personal information from Excel sheets programmatically | Automate Excel data redaction for GDPR/HIPAA compliance
// Developer Intent: Decrypt many protected Excel workbooks, cleanse sensitive content, and re‑save them with a new password in an automated .NET workflow.
// Use Cases: Redact employee SSNs from archived spreadsheets before long‑term storage to meet GDPR or HIPAA rules. | Prepare client financial reports for external audit by stripping personal identifiers and applying updated encryption. | Automate compliance for a batch of vendor‑supplied Excel files, ensuring all files are sanitized and protected with a corporate password.
// AI Prompts: Write C# code using Aspose.Cells that loads all .xlsx files from a folder with an old password, replaces any cell containing "SSN" with "REDACTED", and saves each file with a new password. | Show how to implement per‑file logging of the number of cells redacted and the file path while processing encrypted Excel workbooks in a batch. | Provide an example of exception handling that skips corrupted or incorrectly password‑protected Excel files but continues processing the remaining batch.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchDecryptSanitizeReencrypt
{
    // A .NET console app that scans a folder of password‑protected .xlsx workbooks, opens each with the old password, removes personal data (e.g., comments, SSN values), applies a new password, and saves the cleaned files to an output directory. Includes robust error handling for compliance‑driven bulk processing.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing encrypted Excel files
            string inputFolder = @"C:\InputEncrypted";
            // Folder where re‑encrypted files will be saved
            string outputFolder = @"C:\OutputReencrypted";

            // Old password used to open the encrypted files
            string oldPassword = "OldPassword123";
            // New password to protect the files after sanitization
            string newPassword = "NewSecurePass456";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            try
            {
                // Process each .xlsx file in the input folder
                foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
                {
                    // Verify the file still exists
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook with the old password
                        LoadOptions loadOptions = new LoadOptions
                        {
                            Password = oldPassword
                        };

                        using (Workbook workbook = new Workbook(filePath, loadOptions))
                        {
                            // ---------- Data Sanitization ----------
                            // Remove personal information such as author names in comments
                            workbook.RemovePersonalInformation();

                            // Example sanitization: replace any cell containing "SSN" with "REDACTED"
                            foreach (Worksheet sheet in workbook.Worksheets)
                            {
                                Cells cells = sheet.Cells;
                                int maxRow = cells.MaxDataRow;
                                int maxCol = cells.MaxDataColumn;

                                for (int row = 0; row <= maxRow; row++)
                                {
                                    for (int col = 0; col <= maxCol; col++)
                                    {
                                        Cell cell = cells[row, col];
                                        if (cell.Type == CellValueType.IsString)
                                        {
                                            string text = cell.StringValue;
                                            if (!string.IsNullOrEmpty(text) && text.Contains("SSN"))
                                            {
                                                cell.PutValue("REDACTED");
                                            }
                                        }
                                    }
                                }
                            }

                            // ---------- Re‑encryption ----------
                            // Set the new password for the workbook
                            workbook.Settings.Password = newPassword;

                            // Save the sanitized and re‑encrypted workbook
                            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                            workbook.Save(outputPath, SaveFormat.Xlsx);

                            Console.WriteLine($"Processed and saved: {outputPath}");
                        }
                    }
                    catch (Exception exFile)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {exFile.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
