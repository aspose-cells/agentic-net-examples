// Title: CI Pipeline: Automatically Sign Macro‑Enabled Excel (.xlsm) Workbooks with Aspose.Cells for .NET
// Description: C# sample that loads a macro‑enabled workbook, checks for VBA macros, imports a PFX certificate, creates a DigitalSignature, signs the VBA project and optionally adds an OOXML signature, then saves the signed file and reports the signing status. Ideal for automating macro security in Azure DevOps, GitHub Actions, or any continuous‑integration workflow.
// Keywords: Aspose.Cells | C# digital signature | Excel macro signing | CI/CD | Azure DevOps | GitHub Actions | PFX certificate | VBA project sign | OOXML digital signature | continuous integration | macro security
// Common Searches: how to sign an xlsm file in a CI build using Aspose.Cells | C# code to add a digital signature to a VBA project | automate macro signing with a PFX certificate in Azure DevOps | verify VBA project signature after signing with Aspose.Cells | sign Excel macro workbook in GitHub Actions
// Developer Intent: Apply a digital signature to a macro‑enabled Excel workbook automatically during a CI build.
// Use Cases: Integrate the ExcelMacroSigner into Azure DevOps or GitHub Actions to sign .xlsm artifacts before publishing. | Add a verification step in unit tests that loads the signed workbook and checks IsSigned, IsValidSigned, and IsDigitallySigned flags. | Enforce corporate security policies by applying both VBA project and OOXML signatures to Excel macro files.
// AI Prompts: Generate a PowerShell script that calls ExcelMacroSigner.SignWorkbook using environment variables for paths and logs success or failure. | Create a GitHub Actions workflow step that installs Aspose.Cells, builds the console app, runs the signing method, and fails the job if any signature validation returns false. | Provide enhanced error handling for certificate import failures, including detailed logging of X509KeyStorageFlags and inner exception messages.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace CI_Pipeline
{
    // C# sample that loads a macro‑enabled workbook, checks for VBA macros, imports a PFX certificate, creates a DigitalSignature, signs the VBA project and optionally adds an OOXML signature, then saves the signed file and reports the signing status. Ideal for automating macro security in Azure DevOps, GitHub Actions, or any continuous‑integration workflow.
    public class ExcelMacroSigner
    {
        /// <param name="inputPath">Path to the macro‑enabled workbook (e.g., *.xlsm).</param>
        /// <param name="outputPath">Path where the signed workbook will be saved.</param>
        /// <param name="certificatePath">Path to the PFX certificate file.</param>
        /// <param name="certificatePassword">Password for the PFX certificate.</param>
        public static void SignWorkbook(string inputPath, string outputPath, string certificatePath, string certificatePassword)
        {
            try
            {
                // Verify required files exist
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input workbook not found: {inputPath}");
                    return;
                }
                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                    return;
                }

                // Load the workbook that contains macros
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook actually contains a VBA project
                if (!workbook.HasMacro)
                {
                    Console.WriteLine("The workbook does not contain any macros. No signing performed.");
                    workbook.Save(outputPath, SaveFormat.Xlsm);
                    return;
                }

                // Load the signing certificate (must contain a private key)
                X509Certificate2 certificate = new X509Certificate2();
                certificate.Import(certificatePath, certificatePassword, X509KeyStorageFlags.DefaultKeySet);

                // Create a digital signature object (comments and timestamp are optional)
                DigitalSignature signature = new DigitalSignature(certificate, "CI Pipeline Macro Signing", DateTime.UtcNow);

                // Sign the VBA project
                workbook.VbaProject.Sign(signature);

                // Optional: also add an OOXML digital signature to the workbook itself
                DigitalSignatureCollection dsCollection = new DigitalSignatureCollection();
                dsCollection.Add(signature);
                workbook.SetDigitalSignature(dsCollection);

                // Save the signed workbook (preserve macro format)
                workbook.Save(outputPath, SaveFormat.Xlsm);

                // Verify and report the signing status
                Workbook verify = new Workbook(outputPath);
                Console.WriteLine("VBA Project Signed: " + verify.VbaProject.IsSigned);
                Console.WriteLine("VBA Signature Valid: " + verify.VbaProject.IsValidSigned);
                Console.WriteLine("Workbook Digitally Signed: " + verify.IsDigitallySigned);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during signing: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            // Example usage – replace with actual paths or command‑line arguments
            string inputPath = "sample.xlsm";
            string outputPath = "sample_signed.xlsm";
            string certificatePath = "cert.pfx";
            string certificatePassword = "password";

            ExcelMacroSigner.SignWorkbook(inputPath, outputPath, certificatePath, certificatePassword);
        }
    }
}
