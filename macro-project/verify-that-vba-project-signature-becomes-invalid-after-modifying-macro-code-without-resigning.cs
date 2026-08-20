// Title: C# Example: Detect Invalid VBA Project Signature After Editing a Macro with Aspose.Cells
// Description: Loads a signed .xlsm workbook, checks the VBA project's IsSigned and IsValidSigned flags, appends a comment to the first module, saves to a memory stream, reloads the file, and shows that the project stays signed while the signature becomes invalid.
// Keywords: Aspose.Cells VBA signature | IsValidSigned property | C# modify macro | detect invalid VBA signature | signed .xlsm workbook | VBA project validation | macro code change impact
// Common Searches: how to verify VBA signature with Aspose.Cells .NET | does editing a signed macro invalidate its signature | Aspose.Cells check IsValidSigned after macro change | read and modify VBA module code using Aspose.Cells | C# example for VBA project signature validation
// Developer Intent: Confirm that any alteration to a signed VBA macro renders the existing signature invalid without requiring a new signature.
// Use Cases: Programmatically load a signed macro‑enabled workbook and validate its signature before processing. | Append or modify VBA code, then automatically detect that the signature is no longer valid while the project remains marked as signed. | Integrate signature validation into CI/CD pipelines to reject workbooks whose signed VBA projects have been tampered with.
// AI Prompts: Write C# code using Aspose.Cells that changes a VBA module and asserts workbook.VbaProject.IsValidSigned is false after saving. | Create an NUnit test that loads a signed .xlsm file, modifies a macro line, saves, reloads, and verifies the signature status. | Explain the algorithm Aspose.Cells uses to determine VBA signature validity and which properties should be inspected after editing macro code.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a signed .xlsm workbook, checks the VBA project's IsSigned and IsValidSigned flags, appends a comment to the first module, saves to a memory stream, reloads the file, and shows that the project stays signed while the signature becomes invalid.
    public class VbaSignatureInvalidAfterModificationDemo
    {
        public static void Run()
        {
            // Path to the original workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWithVba.xlsm";

            // Verify that the input file exists
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"File not found: {signedWorkbookPath}");
                return;
            }

            try
            {
                // Load the signed workbook
                Workbook workbook = new Workbook(signedWorkbookPath);

                // Verify that the VBA project is signed and the signature is valid
                Console.WriteLine("Initial state:");
                Console.WriteLine("Is VBA Project Signed: " + workbook.VbaProject.IsSigned);
                Console.WriteLine("Is Signature Valid: " + workbook.VbaProject.IsValidSigned);

                // Ensure there is at least one VBA module to modify
                if (workbook.VbaProject.Modules.Count > 0)
                {
                    // Get the first module
                    VbaModule module = workbook.VbaProject.Modules[0];

                    // Append a comment to the existing code (modifying the macro)
                    module.Codes = module.Codes + "\n' Added comment to invalidate signature";
                }
                else
                {
                    Console.WriteLine("No VBA modules found to modify.");
                    return;
                }

                // Save the modified workbook to a memory stream (preserves macro-enabled format)
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsm);
                    ms.Position = 0; // Reset stream position for reading

                    // Reload the workbook from the modified stream
                    Workbook modifiedWorkbook = new Workbook(ms);

                    // After modification, the signature should be invalid (but still signed)
                    Console.WriteLine("\nAfter modification:");
                    Console.WriteLine("Is VBA Project Signed: " + modifiedWorkbook.VbaProject.IsSigned);
                    Console.WriteLine("Is Signature Valid: " + modifiedWorkbook.VbaProject.IsValidSigned);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VbaSignatureInvalidAfterModificationDemo.Run();
        }
    }
}
