// Title: How to confirm a VBA project signature is invalidated after editing macro code without re‑signing using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# verify VBA project digital signature after changing macro code | detect invalid VBA signature after modifying macro in a signed .xlsm workbook with Aspose.Cells | C# example to read and compare VBA project signature using Aspose.Cells | how to check if editing a VBA module breaks its digital signature in Aspose.Cells .NET | sample code to load a signed macro workbook and test signature integrity with Aspose.Cells
// Tags: Aspose.Cells VBA project signature validation | C# read VBA digital signature Aspose.Cells | modify VBA module without re‑signing Aspose.Cells | compare original and modified VBA signatures .NET | detect broken VBA signature after macro edit

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads a signed .xlsm workbook, uses reflection to obtain the VBA project's digital signature, modifies the first VBA module's code without re‑signing, retrieves the signature again, compares the original and new signatures, reports whether the signature is still valid, and saves the altered workbook as a new file.
class VbaSignatureVerification
{
    static void Main()
    {
        const string inputFile = "SignedMacro.xlsm";
        const string outputFile = "ModifiedMacro.xlsm";

        try
        {
            // Ensure the input workbook exists.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook that contains a VBA project.
            Workbook workbook = new Workbook(inputFile);
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Try to obtain the original digital signature (byte array).
            // Use reflection to stay compatible with different library versions.
            var signatureProp = typeof(VbaProject).GetProperty("Signature");
            byte[] originalSignature = null;
            bool isSigned = false;

            if (signatureProp != null)
            {
                originalSignature = signatureProp.GetValue(vbaProject) as byte[];
                isSigned = originalSignature != null && originalSignature.Length > 0;
            }
            else
            {
                // Fallback: check an IsSigned property if it exists.
                var isSignedProp = typeof(VbaProject).GetProperty("IsSigned");
                if (isSignedProp != null)
                {
                    isSigned = (bool)isSignedProp.GetValue(vbaProject);
                }
            }

            if (!isSigned)
            {
                Console.WriteLine("The VBA project is not signed. Test cannot proceed.");
                return;
            }

            Console.WriteLine($"Original signature length: {originalSignature?.Length ?? 0} bytes.");

            // Modify the code of the first VBA module without re‑signing.
            if (vbaProject.Modules.Count > 0)
            {
                VbaModule module = vbaProject.Modules[0];

                // Use reflection to set the Code property (may not be directly accessible).
                var codeProp = module.GetType().GetProperty("Code");
                if (codeProp != null && codeProp.CanWrite)
                {
                    codeProp.SetValue(module, @"
Sub Test()
    MsgBox ""Macro code has been modified!""
End Sub");
                }
                else
                {
                    Console.WriteLine("Unable to modify VBA module code with the current API.");
                    return;
                }
            }
            else
            {
                Console.WriteLine("No VBA modules found in the project.");
                return;
            }

            // Retrieve the signature after modification (if possible).
            byte[] modifiedSignature = null;
            if (signatureProp != null)
            {
                modifiedSignature = signatureProp.GetValue(vbaProject) as byte[];
            }

            bool signaturesMatch = originalSignature != null &&
                                   originalSignature.SequenceEqual(modifiedSignature ?? Array.Empty<byte>());

            Console.WriteLine(signaturesMatch
                ? "Signature is still valid (unexpected)."
                : "Signature is invalid after macro modification (expected).");

            // Save the modified workbook.
            workbook.Save(outputFile);
            Console.WriteLine($"Modified workbook saved as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
