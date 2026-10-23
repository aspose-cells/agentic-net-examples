// Title: Batch decrypt password‑protected Excel .xlsx files, redact SSN data, and re‑encrypt with a new password using Aspose.Cells for .NET
// Common Searches: How to open password‑protected .xlsx files in C# with Aspose.Cells and replace sensitive text | C# batch process Excel workbooks to redact SSN values and set a new workbook password | Using Aspose.Cells LoadOptions to decrypt multiple Excel files and save them with a different password | Automate sanitization of confidential data in encrypted Excel files with Aspose.Cells .NET | Re‑encrypt Excel files after data masking using Aspose.Cells SaveOptions in C#
// Tags: batch decrypt encrypted Excel workbooks using Aspose.Cells | cell data redaction SSN in Aspose.Cells | programmatic password change for .xlsx with Aspose.Cells | load password‑protected workbook via LoadOptions C# | save sanitized workbook with OoxmlSaveOptions and new password

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // Required for OoxmlSaveOptions

// The C# program iterates through all .xlsx files in a folder, loads each password‑protected workbook using LoadOptions, scans used cells to replace any string containing "SSN" with "REDACTED", assigns a new password via Workbook.Settings, and saves the sanitized file with OoxmlSaveOptions to an output directory while preserving the original file name.
class ExcelBatchProcessor
{
    // Adjust these paths and passwords as needed
    private const string InputFolder = @"C:\InputExcelFiles";
    private const string OutputFolder = @"C:\SanitizedExcelFiles";
    private const string OldPassword = "oldPassword123";
    private const string NewPassword = "newPassword456";

    static void Main()
    {
        try
        {
            // Ensure output directory exists
            Directory.CreateDirectory(OutputFolder);

            // Process each .xlsx file in the input folder
            foreach (string inputFilePath in Directory.GetFiles(InputFolder, "*.xlsx"))
            {
                // Verify that the input file actually exists
                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine($"Input file not found: {inputFilePath}");
                    continue;
                }

                try
                {
                    // Load workbook with the old password
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = OldPassword
                    };
                    Workbook workbook = new Workbook(inputFilePath, loadOptions);

                    // ---- Data sanitization start ----
                    // Example: Replace any cell containing "SSN" with "REDACTED"
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Iterate through used cells only for performance
                        Cells cells = sheet.Cells;
                        int maxRow = cells.MaxDataRow;
                        int maxCol = cells.MaxDataColumn;

                        for (int row = 0; row <= maxRow; row++)
                        {
                            for (int col = 0; col <= maxCol; col++)
                            {
                                Cell cell = cells[row, col];
                                if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("SSN"))
                                {
                                    cell.PutValue("REDACTED");
                                }

                                // Additional sanitization rules can be added here,
                                // e.g., clearing entire columns, masking email addresses, etc.
                            }
                        }
                    }
                    // ---- Data sanitization end ----

                    // Set new password for the workbook
                    workbook.Settings.Password = NewPassword;

                    // Prepare save options (no password property needed)
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

                    // Build output file path (preserve original file name)
                    string outputFilePath = Path.Combine(OutputFolder, Path.GetFileName(inputFilePath));

                    // Save the sanitized workbook with the new password
                    workbook.Save(outputFilePath, saveOptions);

                    Console.WriteLine($"Processed and re‑encrypted: {Path.GetFileName(inputFilePath)}");
                }
                catch (Exception ex)
                {
                    // Log or handle errors per file without stopping the batch
                    Console.WriteLine($"Error processing {Path.GetFileName(inputFilePath)}: {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors that may occur outside the per‑file loop
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
