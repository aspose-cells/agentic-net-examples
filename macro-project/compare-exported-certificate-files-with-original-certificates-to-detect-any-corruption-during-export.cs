// Title: C# console utility to verify that an exported certificate matches the original by comparing SHA‑256 hashes
// AI Prompts: Write a C# console program that accepts two file paths (original and exported certificate), computes their SHA256 hashes with System.Security.Cryptography, and outputs whether the files are identical. | Create a reusable C# method `string GetFileSha256(string path)` that returns the SHA256 hash of any file and throws FileNotFoundException for missing files. | Enhance a C# file‑comparison tool to display both SHA256 values and a clear message indicating a match or possible corruption. | Develop NUnit tests for a certificate export validator that checks hash equality and validates exception handling for nonexistent files.
// Common Searches: how to compare two certificate files for integrity in C# | C# program to check if exported .pfx is corrupted | verify SHA256 hash of a certificate file using .NET | console app to validate exported certificate against original | detect file corruption after exporting certificates with C#
// Tags: compute SHA256 hash of file C# | compare certificate files integrity .NET | file hash validation for exported certificates | C# console utility for certificate integrity check | System.Security.Cryptography SHA256 file comparison

using System;
using System.IO;
using System.Security.Cryptography;

// The example provides a C# console application that computes SHA256 hashes of an original certificate and its exported copy, compares the hashes to detect corruption, prints diagnostic hash values, and includes robust error handling for missing files.
class CertificateExportValidator
{
    // Computes SHA256 hash of a file
    private static string ComputeFileHash(string filePath)
    {
        using (FileStream stream = File.OpenRead(filePath))
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(stream);
            return BitConverter.ToString(hashBytes).Replace("-", "").ToUpperInvariant();
        }
    }

    // Compares two files by their SHA256 hashes and returns true if they are identical
    public static bool AreFilesIdentical(string originalFilePath, string exportedFilePath)
    {
        if (!File.Exists(originalFilePath))
            throw new FileNotFoundException("Original certificate file not found.", originalFilePath);
        if (!File.Exists(exportedFilePath))
            throw new FileNotFoundException("Exported certificate file not found.", exportedFilePath);

        string originalHash = ComputeFileHash(originalFilePath);
        string exportedHash = ComputeFileHash(exportedFilePath);

        // Output hashes for diagnostic purposes
        Console.WriteLine($"Original SHA256: {originalHash}");
        Console.WriteLine($"Exported SHA256: {exportedHash}");

        return string.Equals(originalHash, exportedHash, StringComparison.OrdinalIgnoreCase);
    }

    // Entry point for testing
    static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: CertificateExportValidator <originalCertificatePath> <exportedCertificatePath>");
            return;
        }

        string originalPath = args[0];
        string exportedPath = args[1];

        try
        {
            bool isIdentical = AreFilesIdentical(originalPath, exportedPath);
            if (isIdentical)
                Console.WriteLine("The exported certificate matches the original. No corruption detected.");
            else
                Console.WriteLine("The exported certificate differs from the original. Possible corruption detected.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during validation: {ex.Message}");
        }
    }
}
