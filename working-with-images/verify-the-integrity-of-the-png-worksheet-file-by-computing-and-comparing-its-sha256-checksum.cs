// Title: Validate a worksheet PNG image's integrity by computing and comparing its SHA‑256 checksum in C#
// AI Prompts: Create a C# method that accepts a PNG file path and an expected SHA‑256 hex string, calculates the file's hash, and returns true if the hashes are identical. | Develop a C# console application that reads a worksheet image saved as PNG, computes its SHA‑256 checksum, and outputs a message indicating whether the image passes the integrity check.
// Common Searches: c# how to calculate SHA256 checksum for a PNG exported from Aspose.Cells | verify integrity of worksheet image PNG using SHA256 hash comparison in .NET | compare expected SHA256 value with actual PNG file hash in a C# console program | Aspose.Cells PNG export checksum verification example in C#
// Tags: compute SHA-256 hash for PNG in .NET | worksheet image checksum validation Aspose.Cells | C# file hash comparison utility | PNG export integrity check using SHA256 | SHA256 checksum verification C#

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

// Provides a VerifyPngChecksum method that reads a worksheet PNG file, computes its SHA‑256 hash, and compares it to a supplied hexadecimal hash, returning true on a match; includes a sample console program demonstrating usage.
class PngIntegrityVerifier
{
    // Computes the SHA‑256 hash of the specified file and compares it with the expected hash.
    // Returns true if the hashes match, otherwise false.
    public static bool VerifyPngChecksum(string pngFilePath, string expectedSha256Hex)
    {
        if (!File.Exists(pngFilePath))
            throw new FileNotFoundException("PNG file not found.", pngFilePath);

        // Read the entire PNG file into a byte array.
        byte[] fileBytes = File.ReadAllBytes(pngFilePath);

        // Compute SHA‑256 hash.
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(fileBytes);

            // Convert hash to hexadecimal string for comparison.
            StringBuilder sb = new StringBuilder(hashBytes.Length * 2);
            foreach (byte b in hashBytes)
                sb.AppendFormat("{0:x2}", b);
            string computedHashHex = sb.ToString();

            // Compare computed hash with the expected hash (case‑insensitive).
            return string.Equals(computedHashHex, expectedSha256Hex, StringComparison.OrdinalIgnoreCase);
        }
    }

    // Example usage.
    static void Main()
    {
        string pngPath = @"C:\Temp\WorksheetImage.png";
        string expectedHash = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"; // replace with actual expected hash

        try
        {
            bool isValid = VerifyPngChecksum(pngPath, expectedHash);
            Console.WriteLine(isValid
                ? "PNG file integrity verified. Checksum matches."
                : "PNG file integrity check failed. Checksum does not match.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
