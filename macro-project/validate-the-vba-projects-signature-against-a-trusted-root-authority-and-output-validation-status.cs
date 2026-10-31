// Title: How to verify a VBA project's digital signature against a trusted root certificate using Aspose.Cells for .NET
// Common Searches: C# code to check VBA macro signature in an Excel file with Aspose.Cells | validate Excel VBA project signature using X509Chain in .NET | load VBA project signature bytes from workbook via reflection Aspose.Cells | verify VBA digital signature against a custom root certificate in C# | Aspose.Cells example for VBA signature validation in .NET
// Tags: Aspose.Cells VBA signature extraction | X509Chain trusted root verification | C# VBA project digital signature validation | reflection get SignatureInfo property | load trusted root certificate for chain

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

namespace VbaSignatureValidation
{
    // The sample loads an Excel workbook with Aspose.Cells, detects a VBA project, uses reflection to retrieve the SignatureInfo property, extracts the raw signature bytes, creates an X509Certificate2 from those bytes, loads a trusted root certificate, builds an X509Chain that includes the trusted root, validates the VBA project's digital signature, and outputs whether the signature is valid, handling missing files and errors gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the Excel file containing the VBA project
                string excelPath = "input.xlsx";

                // Verify that the Excel file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Check if the workbook contains a VBA project
                if (workbook.VbaProject != null)
                {
                    // Try to obtain the SignatureInfo property via reflection (available in newer versions)
                    var signatureInfoProp = workbook.VbaProject.GetType().GetProperty("SignatureInfo");
                    if (signatureInfoProp != null)
                    {
                        var signatureInfo = signatureInfoProp.GetValue(workbook.VbaProject);
                        if (signatureInfo != null)
                        {
                            // Retrieve the raw signature bytes
                            var signatureBytesProp = signatureInfo.GetType().GetProperty("Signature");
                            var signatureBytes = signatureBytesProp?.GetValue(signatureInfo) as byte[];

                            if (signatureBytes != null && signatureBytes.Length > 0)
                            {
                                // Load the signature as an X509Certificate2 object
#pragma warning disable SYSLIB0057 // Suppress obsolete constructor warning
                                X509Certificate2 signatureCertificate = new X509Certificate2(signatureBytes);
#pragma warning restore SYSLIB0057

                                // Path to the trusted root authority certificate
                                string trustedRootPath = "trustedRoot.cer";

                                // Verify that the trusted root certificate file exists
                                if (!File.Exists(trustedRootPath))
                                {
                                    Console.WriteLine($"Error: The trusted root certificate \"{trustedRootPath}\" was not found.");
                                    return;
                                }

                                // Load the trusted root certificate
#pragma warning disable SYSLIB0057
                                X509Certificate2 trustedRoot = new X509Certificate2(trustedRootPath);
#pragma warning restore SYSLIB0057

                                // Build a certificate chain to validate the signature against the trusted root
                                X509Chain chain = new X509Chain();
                                chain.ChainPolicy.ExtraStore.Add(trustedRoot);
                                chain.ChainPolicy.VerificationFlags = X509VerificationFlags.AllowUnknownCertificateAuthority;
                                chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;

                                // Perform the validation
                                bool isValid = chain.Build(signatureCertificate);

                                // Output the validation result
                                Console.WriteLine(isValid ? "VBA project signature is valid." : "VBA project signature is invalid.");
                            }
                            else
                            {
                                Console.WriteLine("No digital signature found in the VBA project.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Signature information is not available in this VBA project.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("The current Aspose.Cells version does not expose VBA signature information.");
                    }
                }
                else
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
