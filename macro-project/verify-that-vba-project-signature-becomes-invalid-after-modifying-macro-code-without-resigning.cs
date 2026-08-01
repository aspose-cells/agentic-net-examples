// Title: Check VBA Project Signature Invalidates After Code Change with Aspose.Cells for .NET
// Description: Loads a signed .xlsm workbook, reads VbaProject.IsSigned and IsValidSigned, appends a harmless comment to the first VBA module, saves and reloads the file, and shows that the signature stays present but becomes invalid because the macro was modified without re‑signing.
// Keywords: Aspose.Cells | VBA signature verification | IsValidSigned | .NET | macro tampering detection | signed VBA project | XLSM integrity | VbaProject | re‑sign VBA | code integrity check
// Common Searches: Aspose.Cells check if VBA signature is still valid after editing macro | C# example to read VbaProject.IsSigned and IsValidSigned | detect broken VBA signature in Xlsm file using Aspose.Cells | how to verify VBA project integrity after code change .NET | Aspose.Cells VBA module modification signature invalid
// Developer Intent: Demonstrate that a signed VBA project's digital signature becomes invalid when its macro code is altered without re‑signing.
// Use Cases: Enforce corporate policy by rejecting Excel files whose signed macros have been tampered with. | Automate compliance scans that flag altered VBA code and trigger a re‑signing workflow. | Integrate into a web service that validates uploaded .xlsm files for macro integrity before processing.
// AI Prompts: Generate a C# snippet using Aspose.Cells that loads a signed .xlsm file, modifies a VBA module, and checks IsValidSigned to confirm the signature is broken. | Explain the steps to re‑sign a VBA project after code changes with Aspose.Cells, including certificate selection and signing API calls. | Create NUnit tests that verify IsSigned stays true while IsValidSigned switches to false after appending code to a VBA module.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a signed .xlsm workbook, reads VbaProject.IsSigned and IsValidSigned, appends a harmless comment to the first VBA module, saves and reloads the file, and shows that the signature stays present but becomes invalid because the macro was modified without re‑signing.
    public class VbaSignatureInvalidAfterModificationDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Path to a workbook that already contains a signed VBA project
            string signedWorkbookPath = "SignedVba.xlsm";

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
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project is signed and the signature is initially valid
                Console.WriteLine("Initially - IsSigned: " + vbaProject.IsSigned);
                Console.WriteLine("Initially - IsValidSigned: " + vbaProject.IsValidSigned);

                // Ensure there is at least one module to modify
                if (vbaProject != null && vbaProject.Modules.Count > 0)
                {
                    // Get the first module
                    VbaModule module = vbaProject.Modules[0];

                    // Append a harmless comment to the existing code (modifies the macro)
                    module.Codes = module.Codes + "\n' Modified by Aspose.Cells demo";

                    // Save the modified workbook to a memory stream (macro-enabled format)
                    using (MemoryStream ms = new MemoryStream())
                    {
                        workbook.Save(ms, SaveFormat.Xlsm);
                        ms.Position = 0; // Reset stream position for reading

                        // Reload the workbook from the stream
                        Workbook reloadedWorkbook = new Workbook(ms);
                        VbaProject reloadedVbaProject = reloadedWorkbook.VbaProject;

                        // After modification, the signature should be invalid because we did not re‑sign
                        Console.WriteLine("After modification - IsSigned: " + reloadedVbaProject.IsSigned);
                        Console.WriteLine("After modification - IsValidSigned: " + reloadedVbaProject.IsValidSigned);
                    }
                }
                else
                {
                    Console.WriteLine("The loaded workbook does not contain any VBA modules to modify.");
                }
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine("File not found: " + fnfEx.FileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error processing workbook: " + ex.Message);
            }
        }
    }
}
