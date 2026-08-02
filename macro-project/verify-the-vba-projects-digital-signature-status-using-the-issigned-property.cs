// Title: C# – Detect Digital Signature of VBA Project in Excel (.xlsm) with Aspose.Cells
// Description: Loads an .xlsm workbook using Aspose.Cells for .NET, accesses its VbaProject, and uses the IsSigned and IsValidSigned properties to determine whether the VBA macro is digitally signed and if the signature is valid.
// Keywords: Aspose.Cells | C# VBA signature | VbaProject IsSigned | Excel macro digital signature | IsValidSigned | check signed VBA project | Aspose.Cells .NET | detect signed macro | VBA project verification
// Common Searches: Aspose.Cells check if VBA macro is signed | C# get VBA project signature status | How to verify Excel macro digital signature programmatically | IsSigned property Aspose.Cells example | Determine signed status of .xlsm VBA project
// Developer Intent: Find out whether the VBA project embedded in an Excel workbook is digitally signed and optionally verify its validity.
// Use Cases: Compliance checks before processing workbooks | Security gating – run macros only if signed | Audit logs of macro signing status | Conditional workflow based on signature validity | Automated reporting of signed vs unsigned workbooks
// AI Prompts: Generate C# code using Aspose.Cells to retrieve the signer's certificate details from a signed VBA project. | Show how to handle an unsigned VBA project when loading a workbook with Aspose.Cells, including fallback logic. | Explain steps to validate a VBA project's digital signature and interpret IsSigned vs IsValidSigned in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VbaProject

namespace AsposeCellsExamples
{
    // Loads an .xlsm workbook using Aspose.Cells for .NET, accesses its VbaProject, and uses the IsSigned and IsValidSigned properties to determine whether the VBA macro is digitally signed and if the signature is valid.
    public class VbaProjectSignatureCheckDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string filePath = "sample.xlsm";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook that may contain a VBA project
                Workbook workbook = new Workbook(filePath);

                // Access the VBA project associated with the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Verify whether the VBA project is digitally signed
                if (vbaProject != null && vbaProject.IsSigned)
                {
                    Console.WriteLine("VBA project is signed.");
                    // Optionally, check if the signature is valid
                    Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
                }
                else
                {
                    Console.WriteLine("VBA project is not signed.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
