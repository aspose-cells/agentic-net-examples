// Title: C# – Sign and Verify a VBA Project in an XLSM Workbook with a Self‑Signed Certificate using Aspose.Cells
// Description: Shows how to create a macro‑enabled workbook, generate a 2048‑bit RSA self‑signed certificate, apply it to the workbook’s VbaProject via Aspose.Cells DigitalSignature, save the file, and validate the signature using VbaProject.IsSigned and IsValidSigned.
// Keywords: Aspose.Cells | VBA project signing | self‑signed certificate | C# digital signature | macro‑enabled workbook | VbaProject.Sign | VbaProject.IsSigned | VbaProject.IsValidSigned | Excel macro security | programmatic certificate generation
// Common Searches: sign VBA project Aspose.Cells C# | verify VBA digital signature Aspose.Cells | generate self signed certificate in C# for Excel macro | check IsSigned property Aspose.Cells | programmatically sign macro workbook .NET | Aspose.Cells digital signature example
// Developer Intent: Programmatically sign a VBA project in an XLSM file with a self‑signed certificate and confirm that the signature is recognized.
// Use Cases: Automated CI/CD pipelines that test macro security by signing and verifying VBA projects before release. | Bulk signing of internal macro workbooks with a temporary certificate to meet corporate policy before distribution. | Compliance checks that ensure every VBA project in a workbook is signed and the signature is valid prior to deployment.
// AI Prompts: Provide C# code using Aspose.Cells to sign a VBA project with an existing PFX certificate and handle certificate expiration. | Show how to extract the certificate thumbprint from a signed VbaProject and compare it against a trusted list using Aspose.Cells. | Write error‑handling logic for VbaProject.Sign when the workbook lacks a VBA project or the DigitalSignature is invalid.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSignatureDemo
{
    // Shows how to create a macro‑enabled workbook, generate a 2048‑bit RSA self‑signed certificate, apply it to the workbook’s VbaProject via Aspose.Cells DigitalSignature, save the file, and validate the signature using VbaProject.IsSigned and IsValidSigned.
    class Program
    {
        static void Main()
        {
            // Step 1: Create a new workbook (initially without VBA project)
            Workbook wb = new Workbook();

            // Step 2: Save as macro-enabled workbook to create a VBA project container
            string tempPath = Path.Combine(Path.GetTempPath(), "temp.xlsm");
            wb.Save(tempPath, SaveFormat.Xlsm);

            // Step 3: Load the workbook back – now it contains a VbaProject object
            Workbook macroWb = new Workbook(tempPath);

            // Step 4: Generate a self‑signed certificate (RSA 2048 bits, valid for 1 hour)
            X509Certificate2 certificate;
            using (RSA rsa = RSA.Create(2048))
            {
                var request = new CertificateRequest(
                    new X500DistinguishedName("CN=AsposeSelfSigned"),
                    rsa,
                    HashAlgorithmName.SHA256,
                    RSASignaturePadding.Pkcs1);

                // Create a self‑signed certificate
                certificate = request.CreateSelfSigned(
                    DateTimeOffset.Now,
                    DateTimeOffset.Now.AddHours(1));
            }

            // Step 5: Create a DigitalSignature object using the certificate
            DigitalSignature vbaSignature = new DigitalSignature(
                certificate,
                "Signed by Aspose demo",
                DateTime.Now);

            // Step 6: Sign the VBA project
            VbaProject vbaProject = macroWb.VbaProject;
            if (vbaProject != null)
            {
                vbaProject.Sign(vbaSignature);
            }
            else
            {
                Console.WriteLine("VBA project not found.");
                return;
            }

            // Step 7: Save the signed workbook
            string signedPath = Path.Combine(Environment.CurrentDirectory, "SignedVbaWorkbook.xlsm");
            macroWb.Save(signedPath, SaveFormat.Xlsm);
            Console.WriteLine($"Signed workbook saved to: {signedPath}");

            // Step 8: Reload the workbook to verify the signature
            Workbook verifyWb = new Workbook(signedPath);
            VbaProject verifyProject = verifyWb.VbaProject;

            Console.WriteLine("VBA Project IsSigned: " + verifyProject.IsSigned);
            Console.WriteLine("VBA Project IsValidSigned: " + verifyProject.IsValidSigned);
        }
    }
}
