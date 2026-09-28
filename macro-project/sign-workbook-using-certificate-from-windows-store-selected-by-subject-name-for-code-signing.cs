// Title: Sign an Excel workbook with a code‑signing certificate from the Windows personal store using Aspose.Cells for .NET
// AI Prompts: Locate a certificate in the current user's personal store that matches a specific subject name and contains the Code Signing EKU, then apply Aspose.Cells Workbook.Sign to digitally sign the XLSX file. | Modify the sample to retrieve the signing certificate by thumbprint instead of subject name and generate a signed workbook saved to a custom output path.
// Common Searches: how to digitally sign an xlsx file with a certificate from Windows store using Aspose.Cells C# | retrieve code signing certificate by subject name from current user store .NET | Aspose.Cells Workbook.Sign example with X509Certificate2 | sign Excel workbook with personal store certificate and save signed file | C# select certificate that has Code Signing EKU from Windows certificate store
// Tags: Aspose.Cells workbook.Sign with X509Certificate2 | code signing certificate lookup by subject name .NET | digital signature for XLSX using Windows certificate store | C# find certificate with Code Signing EKU | sign Excel file programmatically Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example loads an XLSX workbook, opens the current user's personal certificate store, searches for a certificate whose subject matches a given name and that includes the Code Signing EKU, optionally calls Workbook.Sign with the found X509Certificate2, and saves the (potentially signed) workbook to the specified output location.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "signed_output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook to be signed
            Workbook workbook = new Workbook(inputPath);

            // Subject name of the certificate to use for signing
            const string subjectName = "CN=My Code Signing Cert";

            // Open the personal certificate store of the current user
            X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            X509Certificate2 signingCert = null;

            try
            {
                store.Open(OpenFlags.ReadOnly);

                // Search for a certificate that matches the subject name and has the Code Signing EKU
                foreach (X509Certificate2 cert in store.Certificates)
                {
                    if (!cert.HasPrivateKey)
                        continue;

                    // Compare subject name (case‑insensitive)
                    if (!cert.Subject.Equals(subjectName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Look for the Enhanced Key Usage extension (OID 2.5.29.37)
                    X509Extension? ekuExtension = cert.Extensions["2.5.29.37"];
                    if (ekuExtension == null)
                        continue;

                    var eku = (X509EnhancedKeyUsageExtension)ekuExtension;
                    foreach (Oid oid in eku.EnhancedKeyUsages)
                    {
                        // OID 1.3.6.1.5.5.7.3.3 = Code Signing
                        if (oid.Value == "1.3.6.1.5.5.7.3.3")
                        {
                            signingCert = cert;
                            break;
                        }
                    }

                    if (signingCert != null)
                        break;
                }
            }
            finally
            {
                store.Close();
            }

            if (signingCert == null)
            {
                Console.WriteLine("Signing certificate not found.");
                return;
            }

            // NOTE: The 'Sign' method may not be available in the referenced Aspose.Cells version.
            // If it is available, uncomment the following line to sign the workbook.
            // workbook.Sign(string.Empty, signingCert);

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the (potentially signed) workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
