// Title: Re‑sign VBA Project in an .xlsm Workbook After Code Changes with Aspose.Cells for .NET
// Description: Shows how to load a macro‑enabled Excel file, edit a VBA module, apply a digital signature using an X509 PFX certificate, save the workbook, and confirm the IsSigned and IsValidSigned status using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# VBA resign | digital signature Excel | macro-enabled workbook signing | X509 certificate .xlsm | VBA project sign | IsValidSigned | Aspose.Cells VbaProject | C# digital signature example | Excel macro security
// Common Searches: how to re‑sign a VBA project after editing with Aspose.Cells | C# code to add a digital signature to an .xlsm file | verify VBA project signature status Aspose.Cells | load X509 certificate and sign VBA modules in Excel | batch resign macro‑enabled workbooks C# | Aspose.Cells example for VBA digital signing
// Developer Intent: Re‑sign the VBA project of a macro‑enabled workbook after making code changes to restore a valid digital signature.
// Use Cases: Update VBA code in an existing .xlsm and apply a new corporate digital signature for compliance. | Automate batch processing of multiple macro‑enabled workbooks: modify modules and resign them with a shared certificate. | Programmatically verify that a workbook’s VBA project remains signed and trusted after modifications.
// AI Prompts: Generate C# code that opens an .xlsm, edits a specific VBA module, and re‑signs the project using a PFX certificate with Aspose.Cells. | Provide best‑practice error handling for loading certificates, signing VBA projects, and checking IsSigned/IsValidSigned properties. | Explain how to loop through a folder of .xlsm files, modify VBA code, resign each workbook, and log any invalid signatures.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaResignDemo
{
    // Shows how to load a macro‑enabled Excel file, edit a VBA module, apply a digital signature using an X509 PFX certificate, save the workbook, and confirm the IsSigned and IsValidSigned status using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the macro-enabled workbook that contains VBA project
                string inputPath = "OriginalWorkbook.xlsm";

                // Verify input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputPath}");
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

                // Example modification: append a comment to the first module (if any)
                if (vbaProject.Modules.Count > 0)
                {
                    VbaModule module = vbaProject.Modules[0];
                    module.Codes += "\r\n' Modified by Aspose.Cells";
                }

                // Load signing certificate (replace with actual path and password)
                string certPath = "MyCertificate.pfx";
                string certPassword = "password";

                // Verify certificate file exists
                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

                // Create a DigitalSignature instance
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,
                    "Resigned after VBA changes",
                    DateTime.Now);

                // Sign the VBA project
                vbaProject.Sign(digitalSignature);

                // Save the workbook (must be saved as macro-enabled format)
                string outputPath = "ResignedWorkbook.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);

                // Verify signature status
                Workbook verifyWorkbook = new Workbook(outputPath);
                Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
                Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
