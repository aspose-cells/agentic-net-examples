// Title: Verify a VBA project's digital signature in an Excel workbook against a trusted root certificate using Aspose.Cells for .NET
// Common Searches: how to check VBA macro signature in an .xlsx file with Aspose.Cells C# | C# code to validate Excel VBA project certificate against a trusted root | using X509Chain to verify VBA project signing certificate in Aspose.Cells | extract VBA project signature bytes from workbook via reflection Aspose.Cells | validate signed macro in Excel workbook with .NET X509Certificate2
// Tags: Aspose.Cells VBA digital signature verification | C# X509Chain chain building for Excel macro certificates | reflection to obtain VBA project signature data | trusted root authority validation for Excel VBA | certificate chain assessment of signed VBA project

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, confirms a VBA project exists, uses reflection to retrieve the project's digital signature bytes, imports them into an X509Certificate2Collection, loads a trusted root certificate, builds an X509Chain for the leaf certificate, ensures the trusted root is present in the chain, and prints whether the VBA project's signature is valid and trusted.
class Program
{
    static void Main()
    {
        // Paths to required files.
        string excelPath = "input.xlsx";
        string trustedRootPath = "trustedRoot.cer";

        // Verify that the input files exist before proceeding.
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
            return;
        }

        if (!File.Exists(trustedRootPath))
        {
            Console.WriteLine($"Error: Trusted root certificate not found at '{trustedRootPath}'.");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(excelPath);

            // Verify that a VBA project exists.
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // Attempt to retrieve the VBA project's digital signature via reflection
            // (the Signature property may not be present in all library versions).
            byte[] signatureData = null;
            var signatureProp = workbook.VbaProject.GetType().GetProperty("Signature");
            if (signatureProp != null)
            {
                signatureData = signatureProp.GetValue(workbook.VbaProject) as byte[];
            }

            if (signatureData == null || signatureData.Length == 0)
            {
                Console.WriteLine("The VBA project is not signed or signature data is unavailable.");
                return;
            }

            // Load the signature into a certificate collection.
            X509Certificate2Collection signedCertificates = new X509Certificate2Collection();
            signedCertificates.Import(signatureData);

            // Load the trusted root certificate.
            X509Certificate2 trustedRoot = new X509Certificate2(trustedRootPath);

            // Build a certificate chain for the leaf certificate (first in the collection).
            X509Chain chain = new X509Chain();
            chain.ChainPolicy.ExtraStore.Add(trustedRoot);
            chain.ChainPolicy.VerificationFlags = X509VerificationFlags.AllowUnknownCertificateAuthority;
            chain.ChainPolicy.RevocationMode = X509RevocationMode.NoCheck;

            X509Certificate2 leafCertificate = signedCertificates[0];
            bool chainIsValid = chain.Build(leafCertificate);

            // Ensure the trusted root appears in the built chain.
            bool trustedRootInChain = false;
            foreach (X509ChainElement element in chain.ChainElements)
            {
                if (element.Certificate.Thumbprint == trustedRoot.Thumbprint)
                {
                    trustedRootInChain = true;
                    break;
                }
            }

            // Output validation result.
            if (chainIsValid && trustedRootInChain)
            {
                Console.WriteLine("VBA project signature is valid and trusted.");
            }
            else
            {
                Console.WriteLine("VBA project signature validation failed.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
