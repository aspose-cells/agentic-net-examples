// Title: Encrypt JSON to a Password‑Protected File with AES‑256 & PBKDF2 in C# (.NET)
// Description: Shows how to create a JSON string in C#, derive a 256‑bit AES key and 128‑bit IV from a password using PBKDF2 (SHA‑256, 100 k iterations) with a random 16‑byte salt, prepend the salt and IV, encrypt the UTF‑8 JSON via AES‑CBC/PKCS7 using CryptoStream, and save the ciphertext to a binary file.
// Keywords: C# AES encryption | JSON encryption .NET | PBKDF2 key derivation C# | CryptoStream AES | password‑protected file | AES‑256 CBC | secure JSON storage | encrypt JSON to file | C# security example
// Common Searches: C# encrypt JSON file with password | AES‑256 encryption of JSON in .NET | PBKDF2 derive key for CryptoStream | store encrypted JSON on disk C# | write salt and IV before ciphertext C#
// Developer Intent: Securely encrypt a generated JSON payload using a password‑derived AES‑256 key and write the encrypted data (including salt and IV) to a file.
// Use Cases: Protect application configuration or user settings by saving them as encrypted JSON on the server. | Securely archive exported reports or analytics data before writing them to disk. | Transmit sensitive JSON payloads over an untrusted network using a shared password for encryption.
// AI Prompts: Generate C# code to decrypt the encryptedData.bin file created by this example using the same password. | Rewrite the encryption routine to use AES‑GCM with built‑in authentication instead of CBC/PKCS7. | Add an HMAC‑SHA256 integrity tag to the encrypted stream and show how to verify it during decryption.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

// Shows how to create a JSON string in C#, derive a 256‑bit AES key and 128‑bit IV from a password using PBKDF2 (SHA‑256, 100 k iterations) with a random 16‑byte salt, prepend the salt and IV, encrypt the UTF‑8 JSON via AES‑CBC/PKCS7 using CryptoStream, and save the ciphertext to a binary file.
class JsonEncryptor
{
    // Generates a sample JSON string (replace with your actual JSON generation logic)
    private static string GenerateJson()
    {
        return @"{ ""Name"": ""John Doe"", ""Age"": 30, ""City"": ""New York"" }";
    }

    // Encrypts the input bytes using AES with a password‑derived key.
    // The output stream contains: [salt][IV][ciphertext]
    private static void EncryptToStream(byte[] plainData, string password, Stream outputStream)
    {
        // Generate a random 16‑byte salt
        byte[] salt = new byte[16];
        using (var rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(salt);
        }

        // Derive key and IV from password and salt (using PBKDF2)
        const int iterations = 100_000; // reasonable security
        using (var keyDerivation = new Rfc2898DeriveBytes(password, salt, iterations, HashAlgorithmName.SHA256))
        {
            byte[] key = keyDerivation.GetBytes(32); // 256‑bit key for AES‑256
            byte[] iv  = keyDerivation.GetBytes(16); // 128‑bit IV

            // Write salt and IV to the beginning of the output (needed for decryption)
            outputStream.Write(salt, 0, salt.Length);
            outputStream.Write(iv, 0, iv.Length);

            // Create AES encryptor
            using (var aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV  = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using (var cryptoStream = new CryptoStream(outputStream, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    cryptoStream.Write(plainData, 0, plainData.Length);
                }
            }
        }
    }

    static void Main()
    {
        // Step 1: Generate JSON content
        string json = GenerateJson();
        byte[] jsonBytes = Encoding.UTF8.GetBytes(json);

        // Step 2: Define password for encryption
        string password = "StrongPassword123!";

        // Step 3: Encrypt JSON and write to a file via a password‑protected stream
        string encryptedFilePath = "encryptedData.bin";
        using (FileStream fileStream = new FileStream(encryptedFilePath, FileMode.Create, FileAccess.Write))
        {
            EncryptToStream(jsonBytes, password, fileStream);
        }

        Console.WriteLine($"JSON data encrypted and saved to '{encryptedFilePath}'.");
    }
}
