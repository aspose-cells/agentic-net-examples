// Title: Validate digital signatures in multiple Excel .xlsx workbooks and generate a CSV summary report with Aspose.Cells for .NET
// AI Prompts: Load each .xlsx file from a directory using Aspose.Cells, invoke DigitalSignatureCollection.ValidateAllSignatures via reflection, and write the file name, total signatures, valid count, invalid count, and overall status to a CSV file. | Add robust error handling to skip missing files, handle unsupported signature features, and log processing errors while building the signature validation report in C#. | Create a summary text file that records each workbook’s signature validation results, using dynamic iteration over the validation result objects returned by Aspose.Cells.
// Common Searches: C# Aspose.Cells batch process to validate digital signatures in all Excel files in a folder | How to generate a CSV report of signature validation results for multiple .xlsx workbooks using Aspose.Cells | Using reflection to access DigitalSignatureCollection in Aspose.Cells when checking workbook signatures | Report overall signature status (AllValid, PartialValid, AllInvalid) for a set of Excel files in .NET
// Tags: Aspose.Cells signature batch validation | C# generate signature summary file | reflection access DigitalSignatureCollection | Excel workbook signature status reporting | validate all signatures Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The C# console application scans a specified folder for .xlsx workbooks, loads each with Aspose.Cells, uses reflection to obtain the DigitalSignatureCollection, validates all signatures, counts total, valid, and invalid signatures, determines an overall status (AllValid, PartialValid, AllInvalid, NoSignature, NotSupported, or Error), and writes these details to a CSV summary file while handling unsupported features and runtime errors.
class SignatureBatchValidator
{
    static void Main()
    {
        // Folder containing the workbooks to be validated
        string inputFolder = @"C:\Workbooks";

        // Path for the summary report
        string reportPath = @"C:\SignatureReport.txt";

        // Ensure the input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }

        // Retrieve all Excel files in the specified folder
        string[] files = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        // Prepare report header
        var reportLines = new List<string>
        {
            "FileName,TotalSignatures,ValidSignatures,InvalidSignatures,OverallStatus"
        };

        foreach (string file in files)
        {
            try
            {
                // Verify the file still exists before loading
                if (!File.Exists(file))
                {
                    Console.WriteLine($"File not found (skipped): {file}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(file);

                // Try to obtain the DigitalSignatureCollection via reflection (avoids compile‑time dependency)
                var sigProp = workbook.GetType().GetProperty("DigitalSignatureCollection");
                if (sigProp == null)
                {
                    // Digital signatures not supported in this version
                    string lineNoSig = $"{Path.GetFileName(file)},0,0,0,NotSupported";
                    reportLines.Add(lineNoSig);
                    continue;
                }

                object signatures = sigProp.GetValue(workbook);
                if (signatures == null)
                {
                    string lineNoSig = $"{Path.GetFileName(file)},0,0,0,NoSignature";
                    reportLines.Add(lineNoSig);
                    continue;
                }

                // Invoke ValidateAllSignatures()
                var validateMethod = signatures.GetType().GetMethod("ValidateAllSignatures");
                if (validateMethod == null)
                {
                    string lineNoSig = $"{Path.GetFileName(file)},0,0,0,NotSupported";
                    reportLines.Add(lineNoSig);
                    continue;
                }

                object validationResults = validateMethod.Invoke(signatures, null);
                if (validationResults == null)
                {
                    string lineNoSig = $"{Path.GetFileName(file)},0,0,0,NoSignature";
                    reportLines.Add(lineNoSig);
                    continue;
                }

                // Use dynamic to iterate over the result collection
                dynamic results = validationResults;
                int total = results.Count;
                int valid = 0;
                int invalid = 0;

                foreach (var result in results)
                {
                    bool isValid = (bool)result.GetType().GetProperty("IsValid")?.GetValue(result);
                    if (isValid)
                        valid++;
                    else
                        invalid++;
                }

                // Determine overall status for the workbook
                string overallStatus = total == 0 ? "NoSignature"
                                     : invalid == 0 ? "AllValid"
                                     : valid == 0 ? "AllInvalid"
                                     : "PartialValid";

                // Add a line to the report
                string line = $"{Path.GetFileName(file)},{total},{valid},{invalid},{overallStatus}";
                reportLines.Add(line);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{file}': {ex.Message}");
                // Record the error in the report
                string errorLine = $"{Path.GetFileName(file)},0,0,0,Error";
                reportLines.Add(errorLine);
            }
        }

        try
        {
            // Ensure the directory for the report exists
            string reportDir = Path.GetDirectoryName(reportPath);
            if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
            {
                Directory.CreateDirectory(reportDir);
            }

            // Write the summary report to the specified file
            File.WriteAllLines(reportPath, reportLines);
            Console.WriteLine($"Signature validation report generated at: {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write report: {ex.Message}");
        }
    }
}
