// Title: Validate an exported .cer certificate against the original file by comparing SHA256 hashes in C#
// Common Searches: C# code to compare two .cer files for integrity using SHA256 | how to detect corruption in exported certificate with .NET hash | verify that exported certificate matches original using file hash in C# | compute SHA256 of a certificate file and compare in .NET | command line tool to check certificate export integrity C#
// Tags: SHA256 hash comparison for .cer files | certificate integrity verification in .NET | compare certificate files using StructuralComparisons | detect export corruption of certificates C# | file hash validation for X509 certificates

using System;
using System.IO;
using System.Security.Cryptography;
using System.Collections;

// The CertificateComparer class provides a static method that validates whether an exported .cer certificate matches the original by computing SHA256 hashes for both files, comparing the resulting byte arrays with StructuralComparisons, and reporting the outcome. It includes a helper ComputeSha256 method and a Main entry point that accepts optional command‑line arguments for the file paths.
class CertificateComparer
{
    // Compare two certificate files by computing their SHA256 hashes.
    // Returns true if the files are identical, false otherwise.
    public static bool CompareCertificates(string originalPath, string exportedPath)
    {
        try
        {
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Original file not found: {originalPath}");
                return false;
            }

            if (!File.Exists(exportedPath))
            {
                Console.WriteLine($"Exported file not found: {exportedPath}");
                return false;
            }

            byte[] originalHash = ComputeSha256(originalPath);
            byte[] exportedHash = ComputeSha256(exportedPath);

            bool areEqual = StructuralComparisons.StructuralEqualityComparer.Equals(originalHash, exportedHash);

            Console.WriteLine(areEqual
                ? "Certificates match. No corruption detected."
                : "Certificates differ! Possible corruption during export.");

            return areEqual;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error comparing certificates: {ex.Message}");
            return false;
        }
    }

    // Compute SHA256 hash of a file.
    private static byte[] ComputeSha256(string filePath)
    {
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha = SHA256.Create())
        {
            return sha.ComputeHash(stream);
        }
    }

    // Example usage.
    static void Main(string[] args)
    {
        try
        {
            // Default file paths; can be overridden by command‑line arguments.
            string originalCertificatePath = @"C:\Certificates\original.cer";
            string exportedCertificatePath = @"C:\Certificates\exported.cer";

            if (args.Length >= 2)
            {
                originalCertificatePath = args[0];
                exportedCertificatePath = args[1];
            }

            CompareCertificates(originalCertificatePath, exportedCertificatePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
