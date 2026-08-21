// Title: Validate VBA Project Digital Signature in an Excel .xlsm Workbook with Aspose.Cells for .NET (C#)
// Description: Loads a signed .xlsm file, confirms a VBA project exists, checks the IsSigned flag, evaluates Workbook.VbaProject.IsValidSigned, and reports verification errors when the signature is invalid or the file has been altered. Includes robust exception handling.
// Keywords: Aspose.Cells VBA signature validation | C# verify Excel macro digital signature | Workbook.VbaProject.IsSigned | Workbook.VbaProject.IsValidSigned example | detect tampered VBA project | Excel macro security check .NET
// Common Searches: how to verify VBA project signature using Aspose.Cells C# | check if Excel macro is signed with Aspose.Cells for .NET | detect invalid VBA digital signature in .xlsm file | Aspose.Cells example for VBA signature validation | C# code to report VBA project verification errors
// Developer Intent: Determine whether a workbook’s VBA project is signed and whether the signature is still valid, then surface any verification errors.
// Use Cases: Pre‑execution security check to ensure macros come from a trusted source. | Batch audit of .xlsm files to flag workbooks with unsigned or tampered VBA projects. | Logging signature validation results for compliance reporting in automated pipelines.
// AI Prompts: Generate C# code that opens an .xlsm file with Aspose.Cells, checks workbook.VbaProject.IsSigned, evaluates IsValidSigned, and prints clear status messages with error handling. | Show how to extract and log detailed verification error information when a VBA project's digital signature fails validation using Aspose.Cells. | Create a reusable method that returns a boolean for VBA signature validity and writes the failure reason to a log file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a signed .xlsm file, confirms a VBA project exists, checks the IsSigned flag, evaluates Workbook.VbaProject.IsValidSigned, and reports verification errors when the signature is invalid or the file has been altered. Includes robust exception handling.
    public class ValidateVbaSignatureDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string filePath = "signedWorkbook.xlsm";

            // Verify that the workbook file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook that contains a VBA project
                Workbook workbook = new Workbook(filePath);

                // Ensure the workbook actually has a VBA project
                if (workbook.VbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Determine whether the VBA project is signed
                if (workbook.VbaProject.IsSigned)
                {
                    Console.WriteLine("VBA project is signed.");

                    // Check if the signature is valid
                    bool isValid = workbook.VbaProject.IsValidSigned;
                    Console.WriteLine("Signature valid: " + isValid);

                    // Report verification error if the signature is not valid
                    if (!isValid)
                    {
                        Console.WriteLine("Verification error: VBA project signature is invalid or the document has been tampered with.");
                    }
                }
                else
                {
                    Console.WriteLine("VBA project is not signed.");
                }
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"File not found error: {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
