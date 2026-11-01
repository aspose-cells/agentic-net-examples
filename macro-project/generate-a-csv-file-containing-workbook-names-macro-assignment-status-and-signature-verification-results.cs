// Title: Generate a CSV report of workbook macro presence and digital signature status using Aspose.Cells for .NET
// Common Searches: C# code to list Excel files with VBA macros and their digital signature results using Aspose.Cells | how to create a CSV file that shows macro detection and signature verification for multiple workbooks in .NET | batch process .xls* files to export macro flag and signature validity with Aspose.Cells | using reflection to read DigitalSignatureCollection in Aspose.Cells C# example | Aspose.Cells macro detection and signature check script for Windows folder
// Tags: Aspose.Cells macro detection CSV export | digital signature verification Aspose.Cells .NET | VBA project enumeration Aspose.Cells | reflection access DigitalSignatureCollection C# | batch workbook analysis Aspose.Cells | UTF‑8 CSV generation C# Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample scans a directory for Excel workbooks, loads each file with Aspose.Cells, determines whether a VBA project exists, attempts to verify the first digital signature via reflection on the DigitalSignatureCollection, and writes a UTF‑8 CSV containing the workbook name, a macro‑presence flag, and the signature verification outcome.
class Program
{
    static void Main()
    {
        try
        {
            // Folder containing the Excel workbooks to analyze
            string folderPath = @"C:\Workbooks";

            // Path of the CSV file to generate
            string csvPath = @"C:\Report\WorkbookReport.csv";

            // Prepare CSV header
            var sb = new StringBuilder();
            sb.AppendLine("WorkbookName,HasMacros,SignatureVerification");

            // Ensure the source folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Source folder does not exist: {folderPath}");
                return;
            }

            // Process each Excel file in the folder
            foreach (string filePath in Directory.GetFiles(folderPath, "*.xls*"))
            {
                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found, skipping: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (Aspose.Cells handles .xls, .xlsx, .xlsm, etc.)
                    var workbook = new Workbook(filePath);

                    // Determine whether the workbook contains VBA macros
                    bool hasMacros = workbook.VbaProject != null && workbook.VbaProject.Modules.Count > 0;

                    // Verify digital signatures (if supported by the library version)
                    string signatureResult = "NotSupported";

                    // Use reflection to access DigitalSignatureCollection safely
                    var sigProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
                    if (sigProp != null)
                    {
                        var signaturesObj = sigProp.GetValue(workbook);
                        if (signaturesObj != null)
                        {
                            // Use dynamic to avoid compile‑time dependency on the type
                            dynamic signatures = signaturesObj;
                            try
                            {
                                if (signatures.Count > 0)
                                {
                                    var signature = signatures[0];
                                    bool isValid = signature.Verify();
                                    signatureResult = isValid ? "Valid" : "Invalid";
                                }
                                else
                                {
                                    signatureResult = "NoSignature";
                                }
                            }
                            catch
                            {
                                // If any member is unavailable, fall back to NotSupported
                                signatureResult = "NotSupported";
                            }
                        }
                    }

                    // Append the result line to the CSV content
                    sb.AppendLine($"{Path.GetFileName(filePath)},{hasMacros},{signatureResult}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(csvPath);
            if (!string.IsNullOrEmpty(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Write the CSV file (UTF‑8 encoding)
            File.WriteAllText(csvPath, sb.ToString(), Encoding.UTF8);
            Console.WriteLine($"Report generated at: {csvPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
