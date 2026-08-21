// Title: Check VBA Project Digital Signature in an Excel .xlsm with Aspose.Cells for .NET
// Description: Loads an .xlsm workbook using Aspose.Cells, accesses its VbaProject, and uses the VbaProject.IsSigned and IsValidSigned properties to determine if the macro project is digitally signed and whether the signature is valid.
// Keywords: Aspose.Cells | VbaProject | IsSigned | IsValidSigned | C# VBA signature | Excel macro digital signature | check signed VBA | xlsm signature verification | Aspose.Cells .NET example
// Common Searches: Aspose.Cells check if VBA project is signed | C# determine VBA macro signature in Excel file | how to verify digital signature of VBA project using Aspose | IsSigned property Aspose.Cells example | validate signed macros in .xlsm with .NET
// Developer Intent: Identify whether the VBA project embedded in an Excel workbook is digitally signed and optionally confirm its validity.
// Use Cases: Security scanning of incoming Excel files to allow only signed macros. | Compliance reporting that logs the signed/unsigned status of VBA projects. | Conditional feature activation when a workbook contains a trusted signed VBA project. | Automated auditing of macro certificates across multiple workbooks.
// AI Prompts: Generate C# code that extracts the signer name and certificate details from a signed VBA project using Aspose.Cells. | Show how to handle workbooks with unsigned VBA projects gracefully, providing fallback logic in Aspose.Cells. | Explain how to validate the full certificate chain of a VBA project's digital signature with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads an .xlsm workbook using Aspose.Cells, accesses its VbaProject, and uses the VbaProject.IsSigned and IsValidSigned properties to determine if the macro project is digitally signed and whether the signature is valid.
    public class VbaProjectSignatureCheckDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            const string filePath = "example.xlsm";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load an Excel workbook that may contain a VBA project
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
                    Console.WriteLine("VBA project is not signed or not present.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
