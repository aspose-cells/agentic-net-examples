// Title: Verify a worksheet PNG image’s integrity by computing and comparing its SHA‑256 checksum in C#
// AI Prompts: Generate a C# console program that reads a PNG file, computes its SHA‑256 hash using System.Security.Cryptography, and prints a pass/fail result based on a supplied expected checksum. | Write a C# method `bool VerifyPngChecksum(string filePath, string expectedHash)` that returns true when the file’s SHA‑256 checksum matches the expected value. | Create a C# snippet that converts the SHA‑256 byte array to a lowercase hexadecimal string for comparison with a known hash.
// Common Searches: how to compute SHA256 hash of a PNG file in C# | C# verify integrity of exported worksheet image using checksum | compare file hash with expected value in a .NET console application | calculate SHA-256 checksum for image files in Aspose.Cells example | C# read binary file and generate SHA256 hex string
// Tags: SHA256 hash calculation for PNG in C# | file checksum verification .NET | worksheet image integrity check using hash | C# console hash validation example | System.Security.Cryptography SHA256 usage

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

// // Reads a worksheet PNG file, computes its SHA‑256 hash with SHA256.Create(), converts the hash to a lowercase hex string, and compares it to a predefined checksum, outputting whether the verification succeeded.
class Program
{
    static void Main()
    {
        // Path to the PNG worksheet file
        string pngPath = "Worksheet.png";

        // Expected SHA-256 checksum (hex string)
        string expectedChecksum = "YOUR_EXPECTED_CHECKSUM_HERE";

        // Load the PNG file into a byte array
        byte[] pngBytes = File.ReadAllBytes(pngPath);

        // Compute SHA-256 hash
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(pngBytes);

            // Convert hash bytes to a hexadecimal string
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hashBytes)
                sb.Append(b.ToString("x2"));
            string actualChecksum = sb.ToString();

            // Compare the computed checksum with the expected value
            if (string.Equals(actualChecksum, expectedChecksum, StringComparison.OrdinalIgnoreCase))
                Console.WriteLine("Checksum verification passed.");
            else
                Console.WriteLine($"Checksum verification failed.\nExpected: {expectedChecksum}\nActual:   {actualChecksum}");
        }
    }
}
