// Title: Re‑sign a VBA project in a macro‑enabled .xlsm workbook after code changes using Aspose.Cells for .NET
// Description: Demonstrates how to load a signed .xlsm file, modify a VBA module, load an X509Certificate2 from a PFX file, create a DigitalSignature, re‑sign the VBA project, save the workbook, and verify the IsSigned and IsValidSigned properties with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | VbaProject | re‑sign VBA | digital signature | X509Certificate2 | PFX certificate | macro‑enabled workbook | xlsm | code modification | signature verification
// Common Searches: re‑sign VBA project Aspose.Cells C# | modify VBA module and preserve digital signature | sign Excel macro workbook with X509Certificate2 | verify VBA project signature after saving | batch re‑sign .xlsm files using Aspose.Cells
// Developer Intent: Update VBA code in a signed macro‑enabled workbook and apply a new digital signature so the file remains valid for distribution.
// Use Cases: Apply code patches to an existing signed .xlsm and re‑sign before release. | Automate re‑signing of multiple macro workbooks after bulk updates. | Confirm that a re‑signed workbook retains a valid signature by checking IsSigned and IsValidSigned.
// AI Prompts: Generate C# code that loads a signed .xlsm, edits a VBA module, and signs the VBA project with a PFX certificate using Aspose.Cells. | Suggest robust error‑handling for missing certificate files, incorrect passwords, or unsigned workbooks during VBA re‑signing. | Explain how to use the IsSigned and IsValidSigned properties to validate a re‑signed VBA project with Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace ReSignVbaProjectDemo
{
    // Demonstrates how to load a signed .xlsm file, modify a VBA module, load an X509Certificate2 from a PFX file, create a DigitalSignature, re‑sign the VBA project, save the workbook, and verify the IsSigned and IsValidSigned properties with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the macro‑enabled workbook that contains a VBA project
                string inputPath = "SignedVba.xlsm";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: '{inputPath}'.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Example modification: append a comment to the first module's code
                if (vbaProject.Modules.Count > 0)
                {
                    VbaModule module = vbaProject.Modules[0];
                    module.Codes += "\r\n' Code modified by ReSignVbaProjectDemo";
                    Console.WriteLine($"Modified module '{module.Name}'.");
                }
                else
                {
                    Console.WriteLine("No VBA modules found to modify.");
                }

                // Load the signing certificate (replace with your actual certificate file and password)
                string certPath = "MyCertificate.pfx";
                string certPassword = "password";

                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: '{certPath}'.");
                    return;
                }

                X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

                // Create a DigitalSignature instance
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,
                    "Re-signed after code change",
                    DateTime.Now);

                // Sign the VBA project
                vbaProject.Sign(digitalSignature);
                Console.WriteLine("VBA project re-signed.");

                // Save the workbook (macro-enabled format)
                string outputPath = "ReSignedVba.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");

                // Verify the signature status by reloading the file
                Workbook verifyWb = new Workbook(outputPath);
                VbaProject verifyVba = verifyWb.VbaProject;

                Console.WriteLine("Verification after reload:");
                Console.WriteLine($"Is Signed: {verifyVba?.IsSigned}");
                Console.WriteLine($"Is Valid Signed: {verifyVba?.IsValidSigned}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
