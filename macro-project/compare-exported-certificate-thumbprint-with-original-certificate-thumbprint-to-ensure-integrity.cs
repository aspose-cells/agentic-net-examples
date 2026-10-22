// Title: C# .NET code to verify a PFX certificate’s integrity by comparing original and re‑imported thumbprints
// Common Searches: C# code to compare thumbprint of a PFX file before and after export | How to check if exported X509Certificate2 matches the original in .NET | Validate certificate integrity by re‑importing a PFX and comparing SHA‑1 thumbprints in C# | Verify that a PFX certificate’s thumbprint remains unchanged after exporting with X509Certificate2 | Sample program for certificate thumbprint verification after export in .NET Core
// Tags: pfx thumbprint verification C# | X509Certificate2 export import integrity check | certificate thumbprint validation .NET | verify certificate integrity after export | reimport PFX thumbprint match

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;

// The example loads a PFX certificate, exports it to a byte array, re‑imports it, retrieves the SHA‑1 thumbprints of both the original and the re‑imported X509Certificate2 instances, and confirms integrity when the thumbprints are identical.
class CertificateIntegrityChecker
{
    // Loads a certificate from a file (PFX) with the given password.
    static X509Certificate2 LoadCertificate(string path, string password)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"Certificate file not found: {path}");

        // The X509KeyStorageFlags ensures the private key is exportable for testing.
        return new X509Certificate2(path, password, X509KeyStorageFlags.Exportable);
    }

    // Exports the certificate (including private key) to a byte array.
    static byte[] ExportCertificate(X509Certificate2 cert)
    {
        // Export as PFX (PKCS#12) with the same password (empty password for simplicity).
        return cert.Export(X509ContentType.Pfx);
    }

    // Computes the thumbprint (SHA-1 hash) of a certificate.
    static string GetThumbprint(X509Certificate2 cert)
    {
        // The Thumbprint property returns the SHA‑1 hash as an uppercase hex string.
        return cert.Thumbprint;
    }

    // Compares the thumbprint of the original certificate with that of the exported‑and‑re‑imported certificate.
    static bool VerifyIntegrity(string certPath, string password)
    {
        try
        {
            // Load the original certificate.
            X509Certificate2 original = LoadCertificate(certPath, password);

            // Export it to a byte array.
            byte[] exportedBytes = ExportCertificate(original);

            // Re‑import the exported bytes as a new certificate instance.
            X509Certificate2 reimported = new X509Certificate2(exportedBytes, password, X509KeyStorageFlags.Exportable);

            // Get thumbprints.
            string originalThumbprint = GetThumbprint(original);
            string reimportedThumbprint = GetThumbprint(reimported);

            // Compare thumbprints (case‑insensitive).
            return string.Equals(originalThumbprint, reimportedThumbprint, StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during integrity verification: {ex.Message}");
            return false;
        }
    }

    // Example usage.
    static void Main()
    {
        string certificatePath = @"C:\Certificates\mycert.pfx";
        string password = "myPassword";

        bool isIntact = VerifyIntegrity(certificatePath, password);

        Console.WriteLine(isIntact
            ? "Certificate integrity verified: thumbprints match."
            : "Certificate integrity check failed: thumbprints do not match.");
    }
}
