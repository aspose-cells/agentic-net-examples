// Title: Validate digital and VBA signatures across multiple Excel workbooks with Aspose.Cells (.NET) – CSV summary
// Description: A C# console utility that scans a specified folder, loads each .xlsx/.xlsm file using Aspose.Cells, determines if the workbook and its VBA project are signed, counts total and valid digital signatures, and writes a CSV file reporting file path, signature presence, and counts.
// Keywords: Aspose.Cells | C# digital signature | Excel batch validation | VBA project signature | CSV report generation | multiple workbooks | signature count | valid signature detection | automation | compliance auditing
// Common Searches: how to validate digital signatures of many Excel files with Aspose.Cells | batch check VBA project signing status in .NET | generate CSV audit of Excel workbook signatures using C# | count valid digital signatures across a folder of spreadsheets | Aspose.Cells example for bulk signature verification
// Developer Intent: Create a batch process that evaluates and logs the signature integrity of Excel workbooks.
// Use Cases: Automated compliance scan of submitted financial reports to ensure each file is properly signed. | Pre‑deployment audit that records macro signing status before publishing workbooks to an internal portal. | CI/CD gate that rejects unsigned or tampered Excel artifacts during build pipelines.
// AI Prompts: Generate a C# method that returns a List of objects with file path, total signatures, and valid‑signature count for a directory of Excel files using Aspose.Cells. | Show how to export the batch signature results to a JSON file instead of CSV in a .NET console app. | Explain the steps to handle password‑protected workbooks while performing bulk signature validation with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using Aspose.Cells.Vba;

namespace BatchSignatureValidation
{
    // Holds validation results for a single workbook
    // A C# console utility that scans a specified folder, loads each .xlsx/.xlsm file using Aspose.Cells, determines if the workbook and its VBA project are signed, counts total and valid digital signatures, and writes a CSV file reporting file path, signature presence, and counts.
    class WorkbookSignatureInfo
    {
        public string FilePath { get; set; }
        public bool IsSigned { get; set; }
        public int SignatureCount { get; set; }
        public int ValidSignatureCount { get; set; }
        public bool VbaSigned { get; set; }
        public bool VbaValidSigned { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Folder containing the workbooks to validate
                string sourceFolder = @"Workbooks";

                // Verify the source folder exists
                if (!Directory.Exists(sourceFolder))
                {
                    Console.WriteLine($"Source folder not found: {sourceFolder}");
                    return;
                }

                // Output CSV report
                string reportPath = @"SignatureSummary.csv";

                // Collect all Excel files (XLSX, XLSM, etc.) in the folder
                string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
                var results = new List<WorkbookSignatureInfo>();

                foreach (string file in workbookFiles)
                {
                    // Ensure the file still exists before attempting to load
                    if (!File.Exists(file))
                    {
                        Console.WriteLine($"File not found, skipping: {file}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        using (Workbook wb = new Workbook(file))
                        {
                            var info = new WorkbookSignatureInfo
                            {
                                FilePath = file,
                                IsSigned = wb.IsDigitallySigned,
                                // VBA project may be null if the workbook has no macros
                                VbaSigned = wb.VbaProject?.IsSigned ?? false,
                                VbaValidSigned = wb.VbaProject?.IsValidSigned ?? false
                            };

                            // If the workbook is digitally signed, retrieve the signature collection
                            if (info.IsSigned)
                            {
                                DigitalSignatureCollection signatures = wb.GetDigitalSignature();
                                if (signatures != null)
                                {
                                    foreach (DigitalSignature sig in signatures)
                                    {
                                        info.SignatureCount++;
                                        if (sig.IsValid)
                                            info.ValidSignatureCount++;
                                    }
                                }
                            }

                            results.Add(info);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{file}': {ex.Message}");
                    }
                }

                // Write the summary report to a CSV file
                using (var writer = new StreamWriter(reportPath))
                {
                    // Header
                    writer.WriteLine("FilePath,IsSigned,SignatureCount,ValidSignatureCount,VbaSigned,VbaValidSigned");
                    // Data rows
                    foreach (var r in results)
                    {
                        writer.WriteLine($"{EscapeCsv(r.FilePath)},{r.IsSigned},{r.SignatureCount},{r.ValidSignatureCount},{r.VbaSigned},{r.VbaValidSigned}");
                    }
                }

                Console.WriteLine($"Signature validation summary written to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Helper to escape commas in file paths for CSV
        private static string EscapeCsv(string field)
        {
            if (field.Contains(",") || field.Contains("\""))
            {
                field = field.Replace("\"", "\"\"");
                return $"\"{field}\"";
            }
            return field;
        }
    }
}
