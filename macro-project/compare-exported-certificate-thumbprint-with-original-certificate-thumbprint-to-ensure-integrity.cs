// Title: C# program to verify certificate integrity by comparing original and exported thumbprints using X509Certificate2
// AI Prompts: Load a .pfx or .cer file with X509Certificate2, export it to a byte array, re‑import the bytes, and compare the Thumbprint values to confirm the export preserved the certificate. | Write C# code that accepts a certificate path and password, extracts the original thumbprint, exports the certificate (including the private key), loads it back, and checks that both thumbprints are identical.
// Common Searches: how to ensure a PFX file remains unchanged after exporting it in C# | C# verify that the thumbprint of an imported certificate matches the original | validate X509Certificate2 export integrity by checking thumbprint equality | check if exported certificate thumbprint matches original using .NET | C# program to confirm certificate export by matching thumbprints
// Tags: X509Certificate2 thumbprint verification after export | PFX integrity check using thumbprint C# | re‑imported certificate thumbprint validation .NET | original and exported certificate thumbprint equality | certificate export integrity with X509Certificate2

using System;
using System.Security.Cryptography.X509Certificates;

// The example loads a certificate file with X509Certificate2, reads its Thumbprint, exports the certificate (including the private key) to a byte array, re‑loads it from that array, retrieves the new Thumbprint, and then compares the two values to ensure the export process did not alter the certificate.
class Program
{
    static void Main(string[] args)
    {
        // Expect two arguments: 
        // 1) Path to the original certificate file (e.g., .pfx or .cer)
        // 2) Password for the certificate (empty string if none)
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <program> <certificatePath> <password>");
            return;
        }

        string certPath = args[0];
        string password = args[1];

        // Load the original certificate
        X509Certificate2 originalCert = new X509Certificate2(certPath, password, X509KeyStorageFlags.Exportable);

        // Get the thumbprint of the original certificate
        string originalThumbprint = originalCert.Thumbprint;
        Console.WriteLine($"Original Thumbprint: {originalThumbprint}");

        // Export the certificate (including private key) to a byte array
        byte[] exportedBytes = originalCert.Export(X509ContentType.Pfx, password);

        // Load a new certificate instance from the exported bytes
        X509Certificate2 importedCert = new X509Certificate2(exportedBytes, password, X509KeyStorageFlags.Exportable);

        // Get the thumbprint of the imported certificate
        string importedThumbprint = importedCert.Thumbprint;
        Console.WriteLine($"Imported Thumbprint: {importedThumbprint}");

        // Compare thumbprints to ensure integrity
        if (string.Equals(originalThumbprint, importedThumbprint, StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("Success: Thumbprints match. Integrity verified.");
        }
        else
        {
            Console.WriteLine("Failure: Thumbprints do not match. Integrity compromised.");
        }
    }
}
