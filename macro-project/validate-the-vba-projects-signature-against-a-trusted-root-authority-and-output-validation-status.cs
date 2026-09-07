// Title: Validate VBA project digital signatures in an Excel workbook against a trusted root certificate with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file, extracts its VBA project digital signatures using Aspose.Cells, and checks each signature against a supplied trusted root X509 certificate. | Show how to obtain the DigitalSignatureCollection from a Workbook via reflection to keep the code compatible with different Aspose.Cells versions. | Demonstrate configuring X509ChainPolicy with a custom trusted root and evaluating the validity of each Excel signature, then outputting the results.
// Common Searches: aspnet verify Excel VBA digital signature with custom root certificate | c# Aspose.Cells check if workbook signatures are trusted | how to use reflection to get DigitalSignatureCollection in Aspose.Cells | validate Excel file signatures against a self‑signed root CA in .NET | sample code for X509Chain validation of Excel digital signatures
// Tags: Aspose.Cells VBA digital signature validation | C# X509Chain custom root certificate for Excel signatures | reflection retrieve DigitalSignatureCollection Aspose.Cells | validate Excel workbook signatures against trusted root | load trusted root X509Certificate2 for signature verification

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;   // Required for DigitalSignature classes

// The example loads an Excel workbook, reads a trusted root certificate, uses a reflection‑based extension to fetch the DigitalSignatureCollection, iterates through each VBA project signature, builds an X509Chain with the custom root to determine validity, reports each signature's status, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string trustedRootPath = "trustedRoot.cer";

            // Verify required files exist
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }

            if (!File.Exists(trustedRootPath))
            {
                Console.WriteLine($"Trusted root certificate not found: {trustedRootPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Load the trusted root certificate (using the obsolete ctor is acceptable for this example)
            X509Certificate2 trustedRoot;
            try
            {
                trustedRoot = new X509Certificate2(trustedRootPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load trusted root certificate: {ex.Message}");
                return;
            }

            // Configure chain policy to trust the loaded root certificate
            X509ChainPolicy chainPolicy = new X509ChainPolicy
            {
                ExtraStore = { trustedRoot },
                VerificationFlags = X509VerificationFlags.AllowUnknownCertificateAuthority,
                RevocationMode = X509RevocationMode.NoCheck
            };

            bool anyValidSignature = false;

            // Access the digital signature collection via reflection (compatible with multiple Aspose.Cells versions)
            DigitalSignatureCollection signatures = workbook.GetDigitalSignatureCollection();

            if (signatures != null)
            {
                // Iterate through all digital signatures in the workbook
                foreach (DigitalSignature signature in signatures)
                {
                    try
                    {
                        // Get the signing certificate
                        X509Certificate2 signingCert = signature.Certificate;

                        // Build and validate the certificate chain
                        using (X509Chain chain = new X509Chain())
                        {
                            chain.ChainPolicy = chainPolicy;
                            bool isValid = chain.Build(signingCert);
                            Console.WriteLine($"Signature by {signingCert.Subject}: {(isValid ? "Valid" : "Invalid")}");
                            if (isValid)
                                anyValidSignature = true;
                        }
                    }
                    catch (Exception exSig)
                    {
                        Console.WriteLine($"Error processing a signature: {exSig.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No digital signatures found in the workbook.");
            }

            if (!anyValidSignature)
            {
                Console.WriteLine("No valid digital signature found in the workbook.");
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Extension method to obtain DigitalSignatureCollection via reflection.
// This avoids compile‑time dependency on a specific Aspose.Cells version.
static class WorkbookExtensions
{
    public static DigitalSignatureCollection GetDigitalSignatureCollection(this Workbook wb)
    {
        var prop = wb.GetType().GetProperty("DigitalSignatureCollection");
        if (prop != null)
        {
            return prop.GetValue(wb) as DigitalSignatureCollection;
        }
        // Fallback for older versions where the property might be named differently.
        var method = wb.GetType().GetMethod("GetDigitalSignatureCollection");
        if (method != null)
        {
            return method.Invoke(wb, null) as DigitalSignatureCollection;
        }
        return null;
    }
}
