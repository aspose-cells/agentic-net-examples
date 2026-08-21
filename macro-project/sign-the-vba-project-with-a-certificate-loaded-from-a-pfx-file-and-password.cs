// Title: C# – Sign an Excel VBA Project with a PFX Certificate using Aspose.Cells
// Description: Demonstrates how to load a macro‑enabled .xlsm workbook, retrieve its VbaProject, import an X509Certificate2 from a PFX file with a password, create a DigitalSignature, sign the VBA project, save the signed workbook, and optionally verify the signature status with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | VBA project signing | PFX certificate | DigitalSignature | VbaProject.Sign | macro-enabled workbook | Excel .xlsm | certificate password | code example
// Common Searches: sign VBA project Aspose.Cells C# | load PFX certificate and sign Excel macro | verify signed VBA project .xlsm | Aspose.Cells VbaProject.Sign sample | C# digital signature for Excel VBA
// Developer Intent: Apply a PFX‑based digital signature to the VBA project of a macro‑enabled Excel file.
// Use Cases: Secure a macro‑enabled workbook before distribution by signing its VBA code. | Automate compliance checks that require VBA projects to be digitally signed. | Integrate VBA signing into a CI/CD pipeline to enforce code‑signing policies.
// AI Prompts: Generate C# code that signs an Excel VBA project with a PFX certificate using Aspose.Cells and includes error handling. | Explain how to confirm that a VBA project is signed and the signature is valid after saving the workbook. | Show how to add VBA project signing to a build script or GitHub Actions workflow with Aspose.Cells for .NET.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigningDemo
{
    // Demonstrates how to load a macro‑enabled .xlsm workbook, retrieve its VbaProject, import an X509Certificate2 from a PFX file with a password, create a DigitalSignature, sign the VBA project, save the signed workbook, and optionally verify the signature status with Aspose.Cells for .NET.
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                VbaProjectSigner.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    public class VbaProjectSigner
    {
        public static void Run()
        {
            // Path to the macro-enabled workbook that contains a VBA project
            string inputPath = "InputWorkbook.xlsm";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }

            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Ensure the workbook actually contains a VBA project
            if (vbaProject != null)
            {
                // Load the signing certificate from a PFX file
                string certPath = "MyCertificate.pfx";
                string certPassword = "certPassword";

                // Verify certificate file exists
                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                X509Certificate2 certificate;
                try
                {
                    certificate = new X509Certificate2(certPath, certPassword);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load certificate: {ex.Message}");
                    return;
                }

                // Create a DigitalSignature instance using the certificate
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,                     // certificate with private key
                    "Signed by Aspose.Cells demo",   // comments / purpose
                    DateTime.Now);                   // signing time

                // Sign the VBA project (feature: VbaProject.Sign)
                vbaProject.Sign(digitalSignature);

                // Save the signed workbook (lifecycle: save)
                string outputPath = "SignedWorkbook.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook signed and saved to: {outputPath}");

                // Optional: Verify the signature after saving
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsm);
                    Workbook verifyWb = new Workbook(ms);
                    Console.WriteLine("VBA Project Signed: " + verifyWb.VbaProject.IsSigned);
                    Console.WriteLine("Signature Valid: " + verifyWb.VbaProject.IsValidSigned);
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
    }
}
