// Title: C# – Sign an Excel VBA project in a .xlsm workbook with a PFX certificate using Aspose.Cells
// Description: Shows how to load a macro‑enabled workbook, retrieve its VbaProject, import an X509Certificate2 from a .pfx file (with password), create a DigitalSignature, sign the VBA project, and save the signed .xlsm file. Includes simple verification of the signature status.
// Keywords: Aspose.Cells | C# VBA signing | sign .xlsm | PFX certificate | X509Certificate2 | DigitalSignature | VbaProject.Sign | macro workbook signing | code signing Excel macros | digital signature Excel
// Common Searches: sign VBA project Aspose.Cells C# | apply PFX certificate to Excel macro workbook | C# code to digitally sign .xlsm file | Aspose.Cells sign macro-enabled workbook | verify VBA project signature programmatically | load .pfx certificate in .NET for Excel signing
// Developer Intent: The developer wants to digitally sign the VBA project of a macro‑enabled workbook using a PFX certificate.
// Use Cases: Apply a corporate code‑signing certificate to protect macros before distributing the workbook. | Automate signing of generated reports that contain embedded macros as part of a CI/CD pipeline. | Ensure end‑users see a trusted signature when opening macro‑enabled workbooks. | Maintain compliance by embedding a digital signature in Excel files that contain VBA code.
// AI Prompts: Generate C# code that loads a .pfx certificate with a password and signs a workbook's VBA project using Aspose.Cells. | Explain how to verify the digital signature of a VBA project after signing it with Aspose.Cells. | Provide error‑handling recommendations for missing VBA project or invalid certificate when signing with Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigningDemo
{
    // Shows how to load a macro‑enabled workbook, retrieve its VbaProject, import an X509Certificate2 from a .pfx file (with password), create a DigitalSignature, sign the VBA project, and save the signed .xlsm file. Includes simple verification of the signature status.
    public class VbaProjectSigner
    {
        /// <param name="inputPath">Path to the source .xlsm workbook.</param>
        /// <param name="outputPath">Path where the signed workbook will be saved.</param>
        /// <param name="pfxPath">Path to the PFX certificate file.</param>
        /// <param name="pfxPassword">Password for the PFX certificate.</param>
        public static void SignVbaProject(string inputPath, string outputPath, string pfxPath, string pfxPassword)
        {
            // Load the workbook that contains a VBA project
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project; it may be null if the workbook has no macros
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Load the certificate (must contain a private key) from the PFX file
            X509Certificate2 certificate = new X509Certificate2(pfxPath, pfxPassword);

            // Create a DigitalSignature instance with the certificate, a comment, and the current time
            DigitalSignature digitalSignature = new DigitalSignature(certificate, "Signed by Aspose.Cells", DateTime.Now);

            // Sign the VBA project
            vbaProject.Sign(digitalSignature);

            // Save the signed workbook in macro-enabled format
            workbook.Save(outputPath, SaveFormat.Xlsm);

            // Optional verification output
            Console.WriteLine($"VBA project signed: {vbaProject.IsSigned}");
            Console.WriteLine($"Signature valid: {vbaProject.IsValidSigned}");
        }

        // Example usage
        public static void Main()
        {
            string inputFile = "SampleWithVba.xlsm";      // source workbook
            string outputFile = "SignedSample.xlsm";      // signed workbook
            string certificateFile = "MyCert.pfx";        // PFX certificate
            string certificatePassword = "certPassword"; // certificate password

            // Ensure the input files exist before attempting to sign
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input workbook not found: {inputFile}");
                return;
            }
            if (!File.Exists(certificateFile))
            {
                Console.WriteLine($"Certificate file not found: {certificateFile}");
                return;
            }

            SignVbaProject(inputFile, outputFile, certificateFile, certificatePassword);
        }
    }
}
