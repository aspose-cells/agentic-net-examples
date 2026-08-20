// Title: C# Batch Validation of Excel Digital Signatures with Aspose.Cells – Console Summary Report
// Description: A C# console utility that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, checks the IsDigitallySigned flag, iterates through the DigitalSignatureCollection, counts valid and invalid signatures, and prints a formatted summary report for every workbook.
// Keywords: Aspose.Cells digital signature verification | .NET Excel signature batch validation | C# workbook IsDigitallySigned | DigitalSignatureCollection count | console report Excel signatures | automated compliance audit Excel | batch Excel signature check | Aspose.Cells API example | Excel file digital signature status
// Common Searches: batch verify digital signatures in Excel files using Aspose.Cells | C# code to list valid and invalid signatures for multiple workbooks | how to generate a summary report of Excel digital signatures | Aspose.Cells IsDigitallySigned example for many files | automate Excel signature validation across a folder
// Developer Intent: Iterate through all .xlsx files in a directory, determine whether each workbook is digitally signed, count valid and invalid signatures, and output a concise console summary.
// Use Cases: Perform a compliance audit of financial spreadsheets by flagging workbooks with missing or invalid digital signatures. | Schedule a nightly job that logs signature validation results to a file for audit‑trail and regulatory reporting. | Integrate the validation loop into a document‑processing pipeline to automatically reject unsigned or tampered workbooks before further analysis. | Provide a quick console‑based health check for a shared repository of Excel reports before distribution.
// AI Prompts: Generate code to export the validation results to a CSV or JSON file instead of printing to the console. | Modify the program to also process .xlsm and .xls files and include a timestamp column in the report. | Add comprehensive exception handling for corrupted or password‑protected workbooks and log errors using a structured logger.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace BatchSignatureValidation
{
    // Represents the validation result for a single workbook
    // A C# console utility that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, checks the IsDigitallySigned flag, iterates through the DigitalSignatureCollection, counts valid and invalid signatures, and prints a formatted summary report for every workbook.
    public class ValidationResult
    {
        public string FileName { get; set; }
        public bool IsSigned { get; set; }
        public int ValidSignatureCount { get; set; }
        public int InvalidSignatureCount { get; set; }
    }

    public class Program
    {
        // Entry point
        public static void Main(string[] args)
        {
            // Folder containing the workbooks to validate
            string folderPath = @"C:\Workbooks";

            // Collect validation results
            List<ValidationResult> results = new List<ValidationResult>();

            // Process each .xlsx file in the folder
            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
            {
                // Load the workbook (uses the provided lifecycle rule)
                Workbook workbook = new Workbook(filePath);

                // Prepare result object
                ValidationResult result = new ValidationResult
                {
                    FileName = Path.GetFileName(filePath),
                    IsSigned = workbook.IsDigitallySigned,
                    ValidSignatureCount = 0,
                    InvalidSignatureCount = 0
                };

                // If the workbook is digitally signed, inspect each signature
                if (result.IsSigned)
                {
                    DigitalSignatureCollection signatures = workbook.GetDigitalSignature();

                    if (signatures != null)
                    {
                        foreach (DigitalSignature signature in signatures)
                        {
                            if (signature.IsValid)
                                result.ValidSignatureCount++;
                            else
                                result.InvalidSignatureCount++;
                        }
                    }
                }

                results.Add(result);
            }

            // Generate a simple summary report to the console
            Console.WriteLine("Batch Digital Signature Validation Report");
            Console.WriteLine("========================================");
            Console.WriteLine($"Processed files: {results.Count}");
            Console.WriteLine();

            foreach (var r in results)
            {
                Console.WriteLine($"File: {r.FileName}");
                Console.WriteLine($"  Signed: {r.IsSigned}");
                if (r.IsSigned)
                {
                    Console.WriteLine($"  Valid Signatures   : {r.ValidSignatureCount}");
                    Console.WriteLine($"  Invalid Signatures : {r.InvalidSignatureCount}");
                }
                Console.WriteLine();
            }
        }
    }
}
