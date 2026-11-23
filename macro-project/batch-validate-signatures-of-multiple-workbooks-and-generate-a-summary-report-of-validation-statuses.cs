// Title: Batch validate digital signatures in multiple Excel workbooks and create a summary report using Aspose.Cells for .NET
// Common Searches: C# code to validate digital signatures in all .xlsx files within a folder using Aspose.Cells | How to generate a signature validation summary workbook for a batch of Excel files with Aspose.Cells | Aspose.Cells digital signature validation when DigitalSignatureCollection API is not available | Create a report of total, valid and invalid signatures for multiple Excel workbooks in C# | Iterate through Excel files and count digital signatures using reflection in Aspose.Cells
// Tags: Aspose.Cells batch digital signature validation | C# generate Excel signature summary report | enumerate .xlsx files for signature analysis | reflection fallback for Aspose.Cells DigitalSignatureCollection | auto‑fit columns in Aspose.Cells report workbook | overall signature status calculation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example scans a given directory (including subfolders) for .xlsx workbooks, loads each file with Aspose.Cells, and uses reflection to access the DigitalSignatureCollection when the API is present. It counts total, valid, and invalid signatures, determines an overall status, and writes these metrics into a new workbook titled "Signature Summary". After processing all files, the columns are auto‑fitted and the report is saved to the specified path.
class BatchSignatureValidator
{
    static void Main()
    {
        try
        {
            // Folder containing the workbooks to validate
            string inputFolder = @"C:\Workbooks";
            // Path for the summary report workbook
            string reportPath = @"C:\SignatureReport.xlsx";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the directory for the report exists
            string reportDir = Path.GetDirectoryName(reportPath);
            if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
                Directory.CreateDirectory(reportDir);

            // Create a new workbook for the summary report
            Workbook reportWorkbook = new Workbook();
            Worksheet sheet = reportWorkbook.Worksheets[0];
            sheet.Name = "Signature Summary";

            // Write header row
            sheet.Cells[0, 0].PutValue("File Name");
            sheet.Cells[0, 1].PutValue("Total Signatures");
            sheet.Cells[0, 2].PutValue("Valid Signatures");
            sheet.Cells[0, 3].PutValue("Invalid Signatures");
            sheet.Cells[0, 4].PutValue("Overall Status");

            int currentRow = 1;

            // Enumerate all Excel files in the input folder (including subfolders)
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.AllDirectories))
            {
                try
                {
                    // Ensure the workbook file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook wb = new Workbook(filePath);

                    // NOTE: Digital signature APIs are available in Aspose.Cells versions that include
                    // the Aspose.Cells.DigitalSignatures namespace. If the current version does not
                    // expose DigitalSignatureCollection, the following placeholder logic will be used.

                    int total = 0;
                    int validCount = 0;
                    int invalidCount = 0;
                    string status = "Digital signature API not available";

                    // Attempt to use the digital signature collection if it exists
                    var signatureCollectionProp = typeof(Workbook).GetProperty("DigitalSignatureCollection");
                    if (signatureCollectionProp != null)
                    {
                        var signatureCollection = signatureCollectionProp.GetValue(wb);
                        if (signatureCollection != null)
                        {
                            var signaturesProp = signatureCollection.GetType().GetProperty("Signatures");
                            var signatures = signaturesProp?.GetValue(signatureCollection) as System.Collections.IEnumerable;
                            if (signatures != null)
                            {
                                foreach (var sig in signatures)
                                {
                                    total++;
                                    // Try to invoke ValidateSignature if available
                                    var validateMethod = sig.GetType().GetMethod("ValidateSignature");
                                    if (validateMethod != null)
                                    {
                                        var result = validateMethod.Invoke(sig, null);
                                        // ValidationResult enum typically has a member named Valid
                                        var isValid = result?.ToString() == "Valid";
                                        if (isValid)
                                            validCount++;
                                        else
                                            invalidCount++;
                                    }
                                }

                                status = total == 0
                                    ? "No Signatures"
                                    : invalidCount == 0
                                        ? "All Valid"
                                        : "Invalid Signatures Present";
                            }
                        }
                    }

                    // Write results to the report sheet
                    sheet.Cells[currentRow, 0].PutValue(Path.GetFileName(filePath));
                    sheet.Cells[currentRow, 1].PutValue(total);
                    sheet.Cells[currentRow, 2].PutValue(validCount);
                    sheet.Cells[currentRow, 3].PutValue(invalidCount);
                    sheet.Cells[currentRow, 4].PutValue(status);

                    currentRow++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            // Auto-fit columns for better readability
            sheet.AutoFitColumns();

            // Save the summary report
            reportWorkbook.Save(reportPath);
            Console.WriteLine($"Signature summary report saved to: {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
