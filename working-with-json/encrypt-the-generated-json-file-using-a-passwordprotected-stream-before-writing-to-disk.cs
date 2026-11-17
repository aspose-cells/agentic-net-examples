// Title: Encrypt a JSON string with a password‑derived AES key using CryptoStream and write to a binary file in C#
// Common Searches: how to encrypt a JSON string with a password using AES in C# | C# example of writing encrypted JSON to a file with CryptoStream | using Rfc2898DeriveBytes to derive key and IV for JSON file encryption | store salt with encrypted JSON data for later decryption in .NET | AES‑CBC encryption of JSON data with password in .NET Core
// Tags: AES-CBC encryption of JSON with CryptoStream | password-derived key generation using Rfc2898DeriveBytes | binary file output of encrypted JSON in .NET | salt prepended to encrypted data for decryption | C# secure JSON file storage

using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;

// The sample converts a JSON string to UTF‑8 bytes, creates a random 16‑byte salt, derives a 256‑bit AES key and a 128‑bit IV with PBKDF2 (Rfc2898DeriveBytes, 100,000 iterations), and encrypts the data using AES‑CBC/PKCS7 via a CryptoStream. The salt is written first, followed by the encrypted JSON, producing a binary file that can be decrypted later with the same password.
class JsonEncryptor
{
    // Encrypts a JSON string using a password‑derived key and writes the encrypted data to a file.
    static void Main()
    {
        // Sample JSON content to be encrypted.
        string jsonContent = "{\"Name\":\"John\",\"Age\":30}";

        // Password used for encryption. In a real scenario, obtain this securely.
        string password = "StrongPassword123!";

        // Path of the output encrypted file.
        string outputPath = "encryptedData.bin";

        // Convert JSON string to bytes.
        byte[] jsonBytes = Encoding.UTF8.GetBytes(jsonContent);

        // Generate a random salt. The same salt must be used for decryption.
        byte[] salt = new byte[16];
        using (var rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(salt);
        }

        // Derive a 256‑bit key and a 128‑bit IV from the password and salt.
        using (var keyDerivation = new Rfc2898DeriveBytes(password, salt, 100_000, HashAlgorithmName.SHA256))
        {
            byte[] key = keyDerivation.GetBytes(32); // 256 bits
            byte[] iv = keyDerivation.GetBytes(16);  // 128 bits

            // Create AES algorithm instance.
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                // Open the output file stream.
                using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    // Write the salt at the beginning of the file (needed for decryption).
                    fileStream.Write(salt, 0, salt.Length);

                    // Create a CryptoStream that encrypts data and writes to the file stream.
                    using (CryptoStream cryptoStream = new CryptoStream(fileStream, aes.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        // Write the JSON bytes into the CryptoStream.
                        cryptoStream.Write(jsonBytes, 0, jsonBytes.Length);
                    } // CryptoStream is closed here, flushing final block.
                } // FileStream is closed here.
            } // AES instance disposed.
        } // Key derivation disposed.

        Console.WriteLine($"JSON data encrypted and saved to '{outputPath}'.");
    }
}
