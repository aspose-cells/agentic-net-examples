// Title: Sign and Verify a VBA Project with a Self‑Signed Certificate using Aspose.Cells for .NET
// Description: Demonstrates how to create a macro‑enabled workbook, generate a self‑signed X509Certificate2, apply a DigitalSignature to the workbook's VbaProject, save the file, and programmatically confirm the signature status with Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA signing | self signed certificate C# | digital signature Excel macro | verify VBA project signature | C# Aspose.Cells example | sign .xlsm file programmatically | VbaProject.Sign Aspose
// Common Searches: how to sign a VBA project with Aspose.Cells | C# self‑signed certificate for Excel macros | validate VBA digital signature after saving | Aspose.Cells example for signing .xlsm | check if VBA project is signed in .NET
// Developer Intent: Programmatically sign a VBA project in an Excel workbook with a self‑signed certificate and verify that the signature is recognized as valid.
// Use Cases: Secure macro code before distribution by applying a digital signature. | Automate integrity checks for macro‑enabled workbooks in CI pipelines. | Batch‑sign multiple .xlsm files in a server‑side .NET application.
// AI Prompts: Generate C# code that creates a self‑signed X509Certificate2 and uses Aspose.Cells to sign a VBA project in a .xlsm file. | Explain step‑by‑step how to verify a VBA project's digital signature after saving the workbook with Aspose.Cells. | Provide guidance on handling certificate expiration and re‑signing VBA projects automatically in a .NET service.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

// Demonstrates how to create a macro‑enabled workbook, generate a self‑signed X509Certificate2, apply a DigitalSignature to the workbook's VbaProject, save the file, and programmatically confirm the signature status with Aspose.Cells for .NET.
class VbaProjectSignDemo
{
    public static void Main()
    {
        // Create a new workbook and save it as a macro‑enabled file to ensure a VBA project exists
        Workbook initialWb = new Workbook();
        string tempPath = Path.Combine(Path.GetTempPath(), "temp.xlsm");
        initialWb.Save(tempPath, SaveFormat.Xlsm);

        // Load the workbook that now contains a VBA project
        Workbook workbook = new Workbook(tempPath);
        VbaProject vbaProject = workbook.VbaProject;
        if (vbaProject == null)
        {
            Console.WriteLine("No VBA project found in the workbook.");
            return;
        }

        // Generate a self‑signed certificate with a private key
        X509Certificate2 certificate;
        using (RSA rsa = RSA.Create(2048))
        {
            var request = new CertificateRequest(
                "CN=AsposeSelfSigned",
                rsa,
                HashAlgorithmName.SHA256,
                RSASignaturePadding.Pkcs1);

            // Add basic constraints (not a CA)
            request.CertificateExtensions.Add(
                new X509BasicConstraintsExtension(false, false, 0, false));

            // Add key usage for digital signatures
            request.CertificateExtensions.Add(
                new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, false));

            // Add subject key identifier
            request.CertificateExtensions.Add(
                new X509SubjectKeyIdentifierExtension(request.PublicKey, false));

            // Create a self‑signed certificate valid for 1 hour
            certificate = request.CreateSelfSigned(DateTimeOffset.Now, DateTimeOffset.Now.AddHours(1));
        }

        // Create a DigitalSignature object from the certificate
        DigitalSignature digitalSignature = new DigitalSignature(
            certificate,
            "Self‑signed VBA signature",
            DateTime.Now);

        // Sign the VBA project with the digital signature
        vbaProject.Sign(digitalSignature);

        // Save the signed workbook to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsm);
            ms.Position = 0; // Reset stream position for reading

            // Reload the workbook to verify the signature
            Workbook verifyWb = new Workbook(ms);
            VbaProject verifyVba = verifyWb.VbaProject;

            Console.WriteLine("Is VBA Project Signed: " + verifyVba.IsSigned);
            Console.WriteLine("Is VBA Signature Valid: " + verifyVba.IsValidSigned);
        }

        // Clean up temporary file
        if (File.Exists(tempPath))
        {
            File.Delete(tempPath);
        }
    }
}
