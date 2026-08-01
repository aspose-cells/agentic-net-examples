// Title: C# – Verify VBA Project IsSigned Status Before and After Digital Signing with Aspose.Cells
// Description: Loads a macro‑enabled workbook, reads the VbaProject.IsSigned and IsValidSigned flags, signs the VBA project with a PFX certificate via Aspose.Cells.DigitalSignature, re‑checks the flags, saves to a memory stream, reloads the file, and confirms that the signature status persists.
// Keywords: Aspose.Cells | C# | .NET | VBA project signing | IsSigned property | IsValidSigned | digital signature | PFX certificate | macro‑enabled workbook | save to memory stream | signature persistence
// Common Searches: How to check if a VBA project is signed using Aspose.Cells .NET | Aspose.Cells C# verify IsSigned after signing a macro workbook | Check IsValidSigned flag before and after VBA digital signature | Persist VBA signature after saving workbook with Aspose.Cells | Sign VBA project programmatically with a PFX certificate
// Developer Intent: Determine and confirm the signing state of a VBA project before and after applying a digital signature, ensuring the status remains after the workbook is saved and reloaded.
// Use Cases: Inspect a macro‑enabled workbook to see if its VBA code is already signed. | Apply a digital signature to a VBA project and validate that IsSigned and IsValidSigned become true. | Save the signed workbook to a stream, reload it, and verify that the signature information is retained.
// AI Prompts: Generate C# code that loads an .xlsm file, displays VbaProject.IsSigned and IsValidSigned, signs the project with a PFX certificate using Aspose.Cells, and re‑checks the flags after saving. | Explain how to create an Aspose.Cells.DigitalSignature from a .pfx file and use it to sign a VBA project. | Suggest robust error‑handling for missing or invalid certificate files when signing a VBA project with Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    // Loads a macro‑enabled workbook, reads the VbaProject.IsSigned and IsValidSigned flags, signs the VBA project with a PFX certificate via Aspose.Cells.DigitalSignature, re‑checks the flags, saves to a memory stream, reloads the file, and confirms that the signature status persists.
    public class VbaProjectSignStatusDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Path to the workbook containing a VBA project
                string inputPath = "sample_with_vba.xlsm"; // adjust as needed
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Display signing status before signing
                Console.WriteLine("Before signing - IsSigned: " + vbaProject.IsSigned);
                Console.WriteLine("Before signing - IsValidSigned: " + vbaProject.IsValidSigned);

                // Path to the certificate used for signing
                string certPath = "certificate.pfx"; // adjust as needed
                string certPassword = "password";    // adjust as needed

                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                // Load the certificate
                X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

                // Create a digital signature instance
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,
                    "Signed by Aspose.Cells demo",
                    DateTime.Now);

                // Sign the VBA project
                vbaProject.Sign(digitalSignature);

                // Verify signing status after signing
                Console.WriteLine("After signing - IsSigned: " + vbaProject.IsSigned);
                Console.WriteLine("After signing - IsValidSigned: " + vbaProject.IsValidSigned);

                // Save the signed workbook to a memory stream and reload to confirm persistence
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsm);
                    ms.Position = 0;

                    Workbook reloadedWorkbook = new Workbook(ms);
                    VbaProject reloadedVba = reloadedWorkbook.VbaProject;

                    Console.WriteLine("After reload - IsSigned: " + reloadedVba.IsSigned);
                    Console.WriteLine("After reload - IsValidSigned: " + reloadedVba.IsValidSigned);
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display them
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
