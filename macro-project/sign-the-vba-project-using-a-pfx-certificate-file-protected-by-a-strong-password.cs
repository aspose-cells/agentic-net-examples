// Title: Digitally Sign a VBA Project in an .xlsm Workbook with a Password‑Protected PFX Certificate using Aspose.Cells for .NET (C#)
// Description: This example shows how to load a macro‑enabled workbook, retrieve its VbaProject, import a PFX certificate protected by a strong password, create a DigitalSignature, sign the VBA project, save the signed file, and verify the signature status with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# sign VBA project | PFX certificate Excel macro signing | password protected certificate digital signature | VbaProject.Sign example | X509Certificate2 load private key | verify VBA signature .xlsm | macro workbook code signing | automated Excel macro signing | CI/CD VBA project signing
// Common Searches: how to sign a VBA project in an .xlsm file using Aspose.Cells | C# code to apply a password protected PFX certificate to an Excel macro workbook | verify digital signature of a VBA project after saving | load X509Certificate2 with private key for Excel macro signing | automate VBA project signing in a CI pipeline
// Developer Intent: The developer needs to digitally sign the VBA project of a macro‑enabled workbook using a password‑protected PFX certificate via Aspose.Cells for .NET.
// Use Cases: Apply a corporate code‑signing certificate to protect VBA macros before distribution. | Automate signing of multiple macro workbooks in a CI/CD pipeline to ensure trusted macros. | Validate the signature after saving to guarantee integrity and authenticity of the VBA code.
// AI Prompts: Generate C# code that signs a VBA project in an .xlsm file using a PFX certificate stored in Azure Key Vault with Aspose.Cells. | Explain how to handle certificate loading failures and fallback to an alternative certificate when signing VBA projects. | Provide a step‑by‑step guide to verify a VBA project's digital signature after saving the workbook with Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace SignVbaProjectDemoApp
{
    // This example shows how to load a macro‑enabled workbook, retrieve its VbaProject, import a PFX certificate protected by a strong password, create a DigitalSignature, sign the VBA project, save the signed file, and verify the signature status with Aspose.Cells for .NET.
    class SignVbaProjectDemo
    {
        static void Main()
        {
            try
            {
                // Path to the macro‑enabled workbook that contains a VBA project
                string workbookPath = "MacroWorkbook.xlsm";

                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Get the VBA project from the workbook
                VbaProject vbaProject = workbook.VbaProject;

                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Path to the PFX certificate file and its password
                string certPath = "MyCertificate.pfx";
                string certPassword = "StrongPassword123!";

                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                // Load the certificate (must contain a private key)
                X509Certificate2 certificate;
                try
                {
                    certificate = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.MachineKeySet);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load certificate: {ex.Message}");
                    return;
                }

                // Create a DigitalSignature instance using the certificate
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,
                    "Signed by Aspose.Cells",
                    DateTime.Now);

                // Sign the VBA project
                vbaProject.Sign(digitalSignature);

                // Save the signed workbook
                string signedPath = "MacroWorkbook_Signed.xlsm";
                workbook.Save(signedPath, SaveFormat.Xlsm);
                Console.WriteLine($"Signed workbook saved to: {signedPath}");

                // Reload the workbook to verify the signature
                Workbook verifyWorkbook = new Workbook(signedPath);
                Console.WriteLine("VBA Project Signed: " + verifyWorkbook.VbaProject.IsSigned);
                Console.WriteLine("Signature Valid: " + verifyWorkbook.VbaProject.IsValidSigned);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
