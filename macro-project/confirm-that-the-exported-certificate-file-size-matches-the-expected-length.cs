// Title: Validate exported .pfx certificate file size against an expected byte count using C#
// Common Searches: c# verify exported pfx certificate file size matches expected bytes | how to compare actual certificate file length with a predefined size in .NET | check if exported certificate file size is correct using File.ReadAllBytes in C#
// Tags: File.ReadAllBytes length validation for certificates | Validate .pfx byte count against expected value | Check existence and size of exported certificate in C# | Actual vs expected certificate file length comparison

using System;
using System.IO;

// The CertificateValidator class offers a static method that confirms a certificate file exists, reads it into a byte array with File.ReadAllBytes, and compares the actual byte length to a supplied expected size, reporting success or mismatch.
class CertificateValidator
{
    /// <param name="certificatePath">Full path to the exported certificate file.</param>
    /// <param name="expectedLength">The expected file size in bytes.</param>
    public static void ValidateCertificateSize(string certificatePath, long expectedLength)
    {
        // Ensure the file exists before attempting to read it.
        if (!File.Exists(certificatePath))
        {
            Console.WriteLine($"Error: Certificate file not found at '{certificatePath}'.");
            return;
        }

        // Load the file into a byte array.
        byte[] fileBytes = File.ReadAllBytes(certificatePath);

        // Compare the actual length with the expected length.
        long actualLength = fileBytes.Length;

        if (actualLength == expectedLength)
        {
            Console.WriteLine($"Success: Certificate size is as expected ({actualLength} bytes).");
        }
        else
        {
            Console.WriteLine($"Failure: Certificate size mismatch. Expected {expectedLength} bytes, but got {actualLength} bytes.");
        }
    }

    // Example usage
    static void Main()
    {
        // Path to the exported certificate file.
        string certPath = @"C:\ExportedCertificates\myCertificate.pfx";

        // Expected file size in bytes (example value).
        long expectedSize = 2048;

        ValidateCertificateSize(certPath, expectedSize);
    }
}
