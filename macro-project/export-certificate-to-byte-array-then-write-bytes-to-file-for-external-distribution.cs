// Title: Export a PFX certificate to a byte array and save it to a file in a C# Aspose.Cells macro project
// Common Searches: C# export X509Certificate2 to byte[] including private key | write certificate byte array to .pfx file with File.WriteAllBytes | Aspose.Cells macro project load workbook and export certificate | load a .pfx certificate with password using X509Certificate2 in .NET | save exported certificate bytes for external distribution in C# | how to use X509Certificate2.Export with X509ContentType.Pfx in a macro
// Tags: certificate export to byte array | write byte array to PFX file using File.WriteAllBytes | Aspose.Cells macro project certificate handling | load X509Certificate2 with password | external distribution of exported certificate | optional workbook load with Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example optionally loads an Excel workbook with Aspose.Cells, then loads a password‑protected PFX certificate, exports the certificate (including its private key) to a byte array, and writes that array to a new .pfx file using File.WriteAllBytes, with comprehensive error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook (optional – only loaded if it exists)
            const string workbookPath = "input.xlsx";

            Workbook workbook = null;
            if (File.Exists(workbookPath))
            {
                try
                {
                    // LoadOptions can include a password if the workbook is protected.
                    var loadOptions = new LoadOptions(LoadFormat.Xlsx);
                    // Uncomment and set the password if required:
                    // loadOptions.Password = "workbookPassword";

                    workbook = new Workbook(workbookPath, loadOptions);
                    Console.WriteLine("Workbook loaded successfully.");
                }
                catch (CellsException ex)
                {
                    // Handle any Aspose.Cells specific errors (e.g., password required)
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Workbook file '{workbookPath}' not found. Skipping workbook load.");
            }

            // Path to the certificate file
            const string certPath = "myCertificate.pfx";
            const string certPassword = "certPassword";

            if (!File.Exists(certPath))
            {
                Console.WriteLine($"Certificate file '{certPath}' not found.");
                return;
            }

            X509Certificate2 certificate;
            try
            {
                // Load the certificate (PFX) with its password
#pragma warning disable SYSLIB0057 // Suppress obsolete warning for demonstration purposes
                certificate = new X509Certificate2(certPath, certPassword);
#pragma warning restore SYSLIB0057
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load certificate: {ex.Message}");
                return;
            }

            // Export the certificate (including the private key) to a byte array
            byte[] certificateBytes = certificate.Export(X509ContentType.Pfx, certPassword);

            // Write the byte array to a file for external distribution
            const string exportPath = "exportedCertificate.pfx";
            try
            {
                File.WriteAllBytes(exportPath, certificateBytes);
                Console.WriteLine("Certificate exported successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write exported certificate: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
