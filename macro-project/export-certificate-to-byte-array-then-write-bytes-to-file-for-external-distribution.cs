// Title: Export a public X509Certificate2 from a PFX to a byte array and write it as a .cer file in C#
// Common Searches: C# code to export public part of a PFX certificate to a .cer file | How to get a byte array from X509Certificate2 and save as DER file in .NET | Saving exported X509Certificate2 bytes to disk using File.WriteAllBytes | Load PFX with Exportable flag and export certificate only in C# | Convert .pfx to .cer programmatically in .NET Core
// Tags: X509Certificate2 export public certificate to byte array | write DER encoded certificate file with File.WriteAllBytes | load PFX with X509KeyStorageFlags.Exportable in .NET | programmatic conversion from .pfx to .cer in C# | certificate export error handling in .NET

using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

// The example loads a PFX file using X509Certificate2 with the Exportable flag, extracts the public certificate as a DER‑encoded byte array via the Export method, ensures the output directory exists, and writes the bytes to a .cer file using File.WriteAllBytes, including handling for cryptographic and I/O exceptions.
class CertificateExporter
{
    static void Main()
    {
        // Path to the PFX certificate and its password
        string certPath = @"C:\Certificates\mycert.pfx";
        string certPassword = "yourPassword";

        // Verify that the certificate file exists before attempting to load it
        if (!File.Exists(certPath))
        {
            Console.WriteLine($"Certificate file not found: {certPath}");
            return;
        }

        try
        {
            // Load the certificate with Exportable flag to allow exporting
            X509Certificate2 certificate = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.Exportable);

            // Export only the public part (DER‑encoded X.509 certificate)
            byte[] certBytes = certificate.Export(X509ContentType.Cert);

            // Define output path for the exported certificate
            string outputFilePath = @"C:\Certificates\mycert_exported.cer";

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Write the exported bytes to the file
            File.WriteAllBytes(outputFilePath, certBytes);

            Console.WriteLine($"Certificate exported successfully to: {outputFilePath}");
        }
        catch (CryptographicException ex)
        {
            Console.WriteLine($"Error loading or exporting certificate: {ex.Message}");
        }
        catch (IOException ex)
        {
            Console.WriteLine($"File I/O error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
