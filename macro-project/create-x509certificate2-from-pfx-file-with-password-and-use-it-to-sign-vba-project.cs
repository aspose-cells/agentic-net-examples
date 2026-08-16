// Title: Sign VBA Project in .xlsm Workbook with a PFX Certificate using Aspose.Cells for .NET
// Description: Demonstrates how to load a macro‑enabled Excel workbook, create an X509Certificate2 from a password‑protected PFX file, build a DigitalSignature, sign the workbook's VBA project, save the signed file, and optionally verify the signature status with Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA signing | C# X509Certificate2 PFX | digital signature Excel macro | sign .xlsm workbook | verify VBA project signature | Aspose.Cells .NET example | certificate based Excel signing
// Common Searches: how to sign a VBA project in an xlsm file using Aspose.Cells | C# load .pfx certificate and sign Excel macro workbook | verify VBA project digital signature after saving | Aspose.Cells sign macro-enabled workbook with certificate | load X509Certificate2 from byte array C#
// Developer Intent: Create an X509Certificate2 from a PFX file and use it to digitally sign the VBA project of a macro‑enabled Excel workbook with Aspose.Cells for .NET.
// Use Cases: Secure a macro‑enabled workbook by signing its VBA project with a corporate certificate. | Programmatically verify that a VBA project is signed and the signature is valid after saving. | Implement robust error handling for missing workbook or certificate files before signing.
// AI Prompts: Generate C# code that loads a password‑protected .pfx file and signs an Excel VBA project using Aspose.Cells. | Show how to check IsSigned and IsValidSigned properties of a VBA project after saving the workbook. | Provide examples of handling file‑not‑found errors when signing a VBA project with a certificate.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigningDemo
{
    // Demonstrates how to load a macro‑enabled Excel workbook, create an X509Certificate2 from a password‑protected PFX file, build a DigitalSignature, sign the workbook's VBA project, save the signed file, and optionally verify the signature status with Aspose.Cells for .NET.
    public class SignVbaProject
    {
        public static void Run()
        {
            try
            {
                // Path to the macro-enabled workbook that contains a VBA project
                string workbookPath = "InputWorkbook.xlsm";

                // Verify workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Ensure the workbook actually contains a VBA project
                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Load the certificate (PFX) with its password
                string certificatePath = "MyCertificate.pfx";
                string certificatePassword = "certPassword";

                // Verify certificate file exists
                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                    return;
                }

                // Load certificate using byte array to avoid obsolete constructor warning
                X509Certificate2 certificate = new X509Certificate2(File.ReadAllBytes(certificatePath), certificatePassword);

                // Create a DigitalSignature object using the certificate
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,                     // certificate to sign with
                    "Signed by Aspose.Cells demo",   // comment/description
                    DateTime.Now);                  // signing time

                // Sign the VBA project
                vbaProject.Sign(digitalSignature);

                // Save the signed workbook as a macro-enabled file
                string signedWorkbookPath = "SignedWorkbook.xlsm";
                workbook.Save(signedWorkbookPath, SaveFormat.Xlsm);

                // Optional: Verify the signature after saving
                Workbook verifyWorkbook = new Workbook(signedWorkbookPath);
                Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
                Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the signing process:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SignVbaProject.Run();
        }
    }
}
