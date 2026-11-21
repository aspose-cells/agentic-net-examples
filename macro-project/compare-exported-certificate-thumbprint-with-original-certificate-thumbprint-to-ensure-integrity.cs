// Title: C# utility to confirm that an exported PFX certificate keeps the same SHA‑1 thumbprint as the source certificate
// Common Searches: how to confirm that an exported PFX retains the original certificate thumbprint in .NET | C# code to compare X509Certificate2 thumbprint before and after using Export method | verify certificate integrity after exporting to a byte array with X509Certificate2 | check if exported certificate matches source thumbprint using .NET Core command line | validate PFX export does not alter SHA‑1 thumbprint in C# application
// Tags: X509Certificate2 export thumbprint verification | certificate thumbprint consistency check .NET | validate PFX export integrity C# | SHA1 thumbprint comparison X509Certificate2 | load certificate from byte array C# | command line certificate integrity tool

using System;
using System.Security.Cryptography.X509Certificates;

// The sample loads an X509Certificate2 from a PFX file, reads its SHA‑1 thumbprint, exports the certificate back to a PFX byte array, reloads it, recomputes the thumbprint, and compares the two values to ensure the export process did not alter the certificate's fingerprint.
class CertificateIntegrityChecker
{
    static void Main(string[] args)
    {
        // Usage: CertificateIntegrityChecker.exe <certificatePath> <password>
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <certificatePath> <password>");
            return;
        }

        string certPath = args[0];
        string password = args[1];

        // Load the original certificate (must be exportable for re‑export)
        X509Certificate2 originalCert = new X509Certificate2(certPath, password, X509KeyStorageFlags.Exportable);

        // Compute thumbprint of the original certificate
        string originalThumbprint = GetThumbprint(originalCert);
        Console.WriteLine($"Original Thumbprint: {originalThumbprint}");

        // Export the certificate to a byte array (PFX format)
        byte[] exportedBytes = originalCert.Export(X509ContentType.Pfx, password);

        // Load the exported certificate from the byte array
        X509Certificate2 exportedCert = new X509Certificate2(exportedBytes, password, X509KeyStorageFlags.Exportable);

        // Compute thumbprint of the exported certificate
        string exportedThumbprint = GetThumbprint(exportedCert);
        Console.WriteLine($"Exported Thumbprint: {exportedThumbprint}");

        // Compare thumbprints to verify integrity
        if (string.Equals(originalThumbprint, exportedThumbprint, StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("Thumbprints match. Integrity verified.");
        }
        else
        {
            Console.WriteLine("Thumbprints do NOT match. Integrity compromised.");
        }
    }

    // Helper method to obtain the SHA‑1 thumbprint of a certificate
    private static string GetThumbprint(X509Certificate2 cert)
    {
        // The Thumbprint property already returns the SHA‑1 hash as a hex string.
        // Normalize to uppercase for a case‑insensitive comparison.
        return cert.Thumbprint?.ToUpperInvariant();
    }
}
