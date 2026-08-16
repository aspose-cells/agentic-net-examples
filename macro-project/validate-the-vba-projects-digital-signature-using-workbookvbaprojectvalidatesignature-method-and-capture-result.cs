// Title: C# – Validate VBA Project Digital Signature in Excel (.xlsm) with Aspose.Cells
// Description: Shows how to load a macro‑enabled workbook, access its VbaProject, check the IsSigned flag, verify the digital signature via IsValidSigned, and save the file while preserving the signature state using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | VbaProject | ValidateSignature | IsValidSigned | C# VBA signature | macro-enabled workbook | digital signature verification | Excel .xlsm | Workbook.VbaProject.ValidateSignature | signed macro detection
// Common Searches: Aspose.Cells validate VBA signature C# | How to check if Excel macro is signed .NET | Workbook.VbaProject.IsValidSigned example | C# code to verify VBA digital signature | Save signed .xlsm after validation Aspose
// Developer Intent: Determine whether a VBA project embedded in an Excel file is signed and whether its digital signature is valid, using Aspose.Cells for .NET.
// Use Cases: Security screening of incoming macro‑enabled files before execution. | Compliance audit to ensure only trusted macros are distributed. | Batch processing of .xlsm files to generate a validation report. | Preserving signature information while programmatically modifying a workbook.
// AI Prompts: Generate C# code that scans a directory of .xlsm files, validates each VBA project's digital signature with Aspose.Cells, and outputs a CSV summary. | Create robust error handling for Workbook.VbaProject.IsValidSigned when the file lacks a VBA project or the signature is corrupted. | Show how to retrieve signer certificate details (subject, issuer, thumbprint) from a signed VBA project using Aspose.Cells. | Write a PowerShell script that invokes a compiled .NET assembly to validate VBA signatures in Excel workbooks.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject

namespace AsposeCellsVbaSignatureValidation
{
    // Shows how to load a macro‑enabled workbook, access its VbaProject, check the IsSigned flag, verify the digital signature via IsValidSigned, and save the file while preserving the signature state using Aspose.Cells for .NET.
    public class ValidateVbaSignature
    {
        public static void Main()
        {
            // Paths for input and output files
            string inputPath = "sample_with_vba.xlsm";   // replace with actual file path
            string outputPath = "validated_output.xlsm";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load a workbook that contains a VBA project (macro-enabled file)
                Workbook workbook = new Workbook(inputPath);

                // Access the VBA project associated with the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Check if the VBA project is signed
                bool isSigned = vbaProject.IsSigned;
                Console.WriteLine("VBA Project Signed: " + isSigned);

                // Validate the digital signature of the VBA project
                // The IsValidSigned property indicates whether the signature is valid
                bool isSignatureValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA Project Signature Valid: " + isSignatureValid);

                // Save the workbook after validation if needed (signature state is preserved)
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
