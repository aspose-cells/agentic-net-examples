// Title: Validate VBA Project Digital Signature in an Excel .xlsm Workbook with Aspose.Cells for .NET
// Description: Loads a signed .xlsm file, confirms the presence of a VBA project, checks if it is digitally signed, and reports the IsValidSigned status to verify trust against a root certificate authority.
// Keywords: Aspose.Cells VBA signature validation | C# verify signed macro workbook | Workbook.VbaProject.IsValidSigned example | Excel macro digital signature .NET | trusted root certificate Aspose.Cells
// Common Searches: how to validate a signed VBA project using Aspose.Cells C# | check if Excel macro workbook is signed and trusted | Aspose.Cells verify VBA project signature against trusted root | C# code to detect invalid VBA signatures in .xlsm files | validate macro digital signature with Aspose.Cells
// Developer Intent: Determine whether the VBA project embedded in an Excel workbook is digitally signed and whether the signature is trusted.
// Use Cases: Enforce security policies by rejecting workbooks with unsigned or tampered macros. | Automate batch verification of macro signatures before importing files into enterprise systems. | Validate uploaded .xlsm files in web applications to ensure only trusted macros are processed.
// AI Prompts: Generate C# code using Aspose.Cells that loads an .xlsm file, detects a VBA project, and validates its digital signature against the system's trusted certificate store. | Explain error‑handling strategies for unsigned VBA projects or invalid signatures, including logging, user notifications, and fallback actions.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a signed .xlsm file, confirms the presence of a VBA project, checks if it is digitally signed, and reports the IsValidSigned status to verify trust against a root certificate authority.
    public class ValidateVbaSignatureDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            string filePath = "signedWorkbook.xlsm";

            // Verify that the workbook file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook that contains a VBA project
            Workbook workbook = new Workbook(filePath);

            // Ensure the workbook has a VBA project before checking its signature
            if (workbook.VbaProject != null && workbook.VbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature is valid: " + workbook.VbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }
        }
    }
}
