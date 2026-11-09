// Title: Validate exported X509 certificate integrity by comparing SHA‑256 hashes in a C# console app
// Common Searches: C# console program to compare SHA256 hash of two certificate files | how to check if exported .pfx file is corrupted using .NET | verify integrity of exported certificate by hashing original and exported files | detect differences between original and exported certificate files in C# | command line tool for certificate file hash verification in .NET
// Tags: certificate SHA256 hash verification C# | exported certificate integrity check .NET | compare two certificate files using SHA256 | detect corrupted certificate file C# | file hash mismatch detection .NET | console app for certificate validation

using System;
using System.IO;
using System.Security.Cryptography;

// A simple C# console utility that accepts the paths of an original and an exported certificate, computes each file’s SHA‑256 hash using System.Security.Cryptography, and reports whether the hashes match, indicating if the exported certificate is corrupted.
class CertificateComparer
{
    static void Main(string[] args)
    {
        // Expect two arguments: path to original certificate and path to exported certificate
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: CertificateComparer <originalFile> <exportedFile>");
            return;
        }

        string originalPath = args[0];
        string exportedPath = args[1];

        bool isCorrupted = IsFileCorrupted(originalPath, exportedPath);
        Console.WriteLine(isCorrupted
            ? "Certificates differ: possible corruption."
            : "Certificates match: no corruption detected.");
    }

    // Compute SHA256 hash of a file
    private static byte[] ComputeHash(string filePath)
    {
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha = SHA256.Create())
        {
            return sha.ComputeHash(stream);
        }
    }

    // Compare two files by their hash values; returns true if they differ (i.e., corrupted)
    private static bool IsFileCorrupted(string originalFile, string exportedFile)
    {
        if (!File.Exists(originalFile) || !File.Exists(exportedFile))
        {
            Console.WriteLine("One or both files do not exist.");
            return true;
        }

        byte[] originalHash = ComputeHash(originalFile);
        byte[] exportedHash = ComputeHash(exportedFile);

        // If hash lengths differ, files are different
        if (originalHash.Length != exportedHash.Length)
            return true;

        // Compare each byte of the hash
        for (int i = 0; i < originalHash.Length; i++)
        {
            if (originalHash[i] != exportedHash[i])
                return true;
        }

        // Hashes match, files are identical
        return false;
    }
}
