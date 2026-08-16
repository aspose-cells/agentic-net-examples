// Title: C# Example: Verify VBA Project IsSigned After Adding a Digital Signature with Aspose.Cells for .NET
// Description: Creates a macro‑enabled workbook, loads a .pfx certificate, signs the workbook's VBA project using Aspose.Cells.DigitalSignatures, saves to a memory stream, reloads the file, and reads VbaProject.IsSigned and VbaProject.IsValidSigned to confirm the signature status.
// Keywords: Aspose.Cells | C# VBA digital signature | VbaProject.IsSigned | VbaProject.IsValidSigned | macro‑enabled workbook | X509Certificate2 signing | programmatic VBA signing | .NET Excel automation | GitHub Aspose.Cells example
// Common Searches: how to check if a VBA project is signed using Aspose.Cells .NET | Aspose.Cells sample code to sign VBA project and verify IsSigned | C# verify VBA digital signature after signing | VbaProject.IsValidSigned returns false after signing | load .pfx certificate and sign Excel macro project with Aspose
// Developer Intent: Confirm that a VBA project has been digitally signed and that the signature is valid after programmatic signing with Aspose.Cells.
// Use Cases: Generate a new .xlsm workbook, apply a digital signature to its VBA project, and read IsSigned to ensure the signing succeeded. | Load an existing .pfx certificate, sign the VBA project of any workbook, then check both IsSigned and IsValidSigned for validation. | Handle missing certificate or missing VBA project gracefully while still reporting the signing outcome.
// AI Prompts: Write C# code using Aspose.Cells to sign a VBA project with a .pfx certificate and then display VbaProject.IsSigned and IsValidSigned. | Explain error handling when the certificate file is absent while attempting to sign a VBA project with Aspose.Cells. | Show a step‑by‑step example that creates a macro‑enabled workbook, signs its VBA project, saves to a stream, reloads, and verifies the signature status.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

namespace AsposeCellsExamples
{
    // Creates a macro‑enabled workbook, loads a .pfx certificate, signs the workbook's VBA project using Aspose.Cells.DigitalSignatures, saves to a memory stream, reloads the file, and reads VbaProject.IsSigned and VbaProject.IsValidSigned to confirm the signature status.
    class CheckVbaSignature
    {
        static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (contains a VBA project)
            Workbook wb = new Workbook();

            // Ensure the VBA project exists by saving as macro-enabled and reloading
            using (MemoryStream tempStream = new MemoryStream())
            {
                wb.Save(tempStream, SaveFormat.Xlsm);
                tempStream.Position = 0;
                wb = new Workbook(tempStream);
            }

            // Load the signing certificate if the file exists
            const string certPath = "YourCertificate.pfx";
            const string certPassword = "password";
            X509Certificate2? certificate = null;

            if (File.Exists(certPath))
            {
                try
                {
#pragma warning disable SYSLIB0057 // Suppress obsolete warning for demo purposes
                    certificate = new X509Certificate2(certPath, certPassword);
#pragma warning restore SYSLIB0057
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load certificate: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Certificate file not found: {certPath}");
            }

            // Sign the VBA project if a certificate was loaded
            VbaProject? vbaProject = wb.VbaProject;
            if (vbaProject != null && certificate != null)
            {
                DigitalSignature vbaSignature = new DigitalSignature(certificate, "VBA Signing", DateTime.Now);
                vbaProject.Sign(vbaSignature);
            }
            else
            {
                Console.WriteLine("VBA project not signed (missing certificate or VBA project).");
            }

            // Save the signed workbook to a memory stream and verify the signature
            using (MemoryStream signedStream = new MemoryStream())
            {
                wb.Save(signedStream, SaveFormat.Xlsm);
                signedStream.Position = 0;

                Workbook verifyWb = new Workbook(signedStream);
                VbaProject? verifyVba = verifyWb.VbaProject;

                if (verifyVba != null)
                {
                    Console.WriteLine("VBA Project IsSigned: " + verifyVba.IsSigned);
                    Console.WriteLine("VBA Project IsValidSigned: " + verifyVba.IsValidSigned);
                }
                else
                {
                    Console.WriteLine("No VBA project found in the saved workbook.");
                }
            }
        }
    }
}
