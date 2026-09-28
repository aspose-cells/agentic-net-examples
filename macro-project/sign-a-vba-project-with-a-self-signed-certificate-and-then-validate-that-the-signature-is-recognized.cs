// Title: Sign a VBA project in an Xlsm workbook with a self‑signed certificate and verify the signature using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a self‑signed ECDSA X509 certificate, adds a VBA module to a macro‑enabled workbook, signs the VBA project with Aspose.Cells, and saves the file as .xlsm. | Write C# to load a signed .xlsm workbook with Aspose.Cells and programmatically confirm the VBA project signature using VerifySignature or the IsSigned property, handling any exceptions. | Implement fallback signing logic in C# that first calls VbaProject.Sign(string certPath, string password) and, if it fails, uses VbaProject.Sign(X509Certificate2 cert).
// Common Searches: asp.net sign vba project in xlsm using self signed certificate aspose.cells | c# verify vba macro signature after saving workbook with aspose.cells | how to use Aspose.Cells VbaProject.Sign overload certificate path password | create self signed certificate for Excel macro signing c# | check if VBA project is signed in loaded workbook Aspose.Cells
// Tags: VbaProject.Sign with certificate file Aspose.Cells | VbaProject.Sign using X509Certificate2 overload | verify VBA project signature Aspose.Cells | self‑signed certificate generation C# for Excel macros | macro‑enabled workbook signature verification .xlsm

using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

namespace AsposeCellsVbaSigning
{
    // The example creates a self‑signed ECDSA certificate (if missing), builds a macro‑enabled workbook, adds a VBA module, signs the VBA project using either the certificate‑file overload or an X509Certificate2 instance, saves the workbook as .xlsm, reloads it, and validates the signature via VerifySignature or the IsSigned property, outputting the verification result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for the self‑signed certificate and the output workbook
                string certPath = "SelfSignedCert.pfx";
                string certPassword = "password";
                string signedFile = "SignedWorkbook.xlsm";

                // Create a self‑signed certificate if it does not exist
                if (!File.Exists(certPath))
                {
                    using (ECDsa ecdsa = ECDsa.Create())
                    {
                        var req = new CertificateRequest(
                            "cn=SelfSignedVbaCert",
                            ecdsa,
                            HashAlgorithmName.SHA256);

                        using (X509Certificate2 cert = req.CreateSelfSigned(
                            DateTimeOffset.Now.AddDays(-1),
                            DateTimeOffset.Now.AddYears(1)))
                        {
                            byte[] pfxBytes = cert.Export(X509ContentType.Pfx, certPassword);
                            File.WriteAllBytes(certPath, pfxBytes);
                        }
                    }
                }

                // Create a new macro‑enabled workbook
                var workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";

                // Work with VBA project using dynamic to stay compatible with different library versions
                dynamic vbaProject = workbook.VbaProject;

                // Add a VBA module
                dynamic vbaModule = vbaProject.Modules.Add("Module1");
                vbaModule.Codes = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";

                // Sign the VBA project (try both overloads)
                try
                {
                    // Preferred overload (certificate file + password)
                    vbaProject.Sign(certPath, certPassword);
                }
                catch
                {
                    // Fallback overload (X509Certificate2 instance)
                    var cert = new X509Certificate2(certPath, certPassword);
                    vbaProject.Sign(cert);
                }

                // Save the signed workbook
                workbook.Save(signedFile, SaveFormat.Xlsm);

                // Reload the workbook and verify the signature
                if (File.Exists(signedFile))
                {
                    var loadedWorkbook = new Workbook(signedFile);
                    dynamic loadedVba = loadedWorkbook.VbaProject;
                    bool isSignatureValid = false;

                    try
                    {
                        // Preferred verification method
                        isSignatureValid = loadedVba.VerifySignature();
                    }
                    catch
                    {
                        try
                        {
                            // Fallback property (if available)
                            isSignatureValid = loadedVba.IsSigned;
                        }
                        catch
                        {
                            // If neither is available, assume verification is not supported
                            isSignatureValid = false;
                        }
                    }

                    Console.WriteLine($"VBA project signature verified: {isSignatureValid}");
                }
                else
                {
                    Console.WriteLine("Signed workbook file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
