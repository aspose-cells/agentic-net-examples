// Title: Validate VBA Project Digital Signature and Detect Errors with Aspose.Cells for .NET (C#)
// Description: Loads an .xlsm workbook, determines whether its VBA project is signed, validates the digital signature using Workbook.VbaProject.IsValidSigned, and outputs status or error messages when the signature is invalid or the file has been altered.
// Keywords: Aspose.Cells | C# VBA signature validation | Workbook.VbaProject.IsSigned | Workbook.VbaProject.IsValidSigned | Excel macro digital signature | detect tampered VBA | signed .xlsm verification | macro security Aspose
// Common Searches: How to check VBA project signature with Aspose.Cells C# | Validate digital signature of Excel macro using Aspose.Cells | Detect tampered VBA macros in .xlsm files .NET | Aspose.Cells IsSigned vs IsValidSigned | Report VBA signature errors in C#
// Developer Intent: The developer wants to determine whether a VBA project in an Excel workbook is signed, verify the signature’s validity, and report any verification errors.
// Use Cases: Prevent processing of workbooks whose VBA macros have been tampered with by validating the digital signature first. | Log signing status and validation results for compliance auditing of macro‑enabled Excel files. | Display user‑friendly messages or halt execution when an unsigned or invalidly signed VBA project is detected. | Integrate signature checks into automated security scans of incoming .xlsm files.
// AI Prompts: Generate C# code using Aspose.Cells that loads a workbook, checks Workbook.VbaProject.IsSigned, validates the signature with IsValidSigned, and outputs detailed error information if the signature is invalid. | Explain the difference between Workbook.VbaProject.IsSigned and Workbook.VbaProject.IsValidSigned properties and show how to handle unsigned, valid, and invalid VBA projects in a .NET application. | Provide best practices for reporting VBA signature verification errors and integrating signature checks into a macro processing workflow with Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an .xlsm workbook, determines whether its VBA project is signed, validates the digital signature using Workbook.VbaProject.IsValidSigned, and outputs status or error messages when the signature is invalid or the file has been altered.
class ValidateVbaSignature
{
    static void Main()
    {
        // Load the workbook that contains a VBA project.
        Workbook workbook = new Workbook("SignedWorkbook.xlsm");

        // Determine whether the VBA project is signed.
        if (workbook.VbaProject.IsSigned)
        {
            Console.WriteLine("VBA project is signed.");

            // Verify the validity of the signature.
            bool isSignatureValid = workbook.VbaProject.IsValidSigned;
            Console.WriteLine("Signature valid: " + isSignatureValid);

            // Report any verification errors.
            if (!isSignatureValid)
            {
                Console.WriteLine("Error: The VBA project signature is invalid or the document has been tampered with.");
            }
        }
        else
        {
            Console.WriteLine("VBA project is not signed.");
        }
    }
}
