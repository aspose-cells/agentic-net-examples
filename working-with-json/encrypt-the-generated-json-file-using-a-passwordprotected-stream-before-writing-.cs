using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

class Program
{
    // Password used to protect the JSON data
    private const string Password = "StrongPassword123!";

    // Salt for key derivation (should be stored with the encrypted file)
    private static readonly byte[] Salt = Encoding.UTF8.GetBytes("UniqueSaltValue");

    static void Main()
    {
        // Sample data to be serialized to JSON
        var data = new
        {
            Id = 1,
            Name = "Sample Item",
            Timestamp = DateTime.UtcNow
        };

        // Serialize the object to JSON string
        string jsonString = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });

        // Convert JSON string to bytes
        byte[] plainBytes = Encoding.UTF8.GetBytes(jsonString);

        // Encrypt the JSON bytes and write to a file
        string encryptedFilePath = "encryptedData.bin";
        using (FileStream fileStream = new FileStream(encryptedFilePath, FileMode.Create, FileAccess.Write))
        {
            EncryptAndWrite(plainBytes, fileStream);
        }

        Console.WriteLine($"Encrypted JSON saved to '{encryptedFilePath}'.");
    }

    // Encrypts the input data using AES with a password-derived key and writes to the provided stream.
    private static void EncryptAndWrite(byte[] data, Stream outputStream)
    {
        // Derive a 256‑bit key and a 128‑bit IV from the password and salt
        using (var keyDerivation = new Rfc2898DeriveBytes(Password, Salt, 100_000, HashAlgorithmName.SHA256))
        {
            byte[] key = keyDerivation.GetBytes(32); // 256 bits
            byte[] iv = keyDerivation.GetBytes(16);  // 128 bits

            // Create AES encryptor
            using (Aes aes = Aes.Create())
            {
                aes.Key = key;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                // Write the salt at the beginning of the file so it can be used for decryption
                outputStream.Write(Salt, 0, Salt.Length);

                // Create a CryptoStream that encrypts data as it is written
                using (CryptoStream cryptoStream = new CryptoStream(outputStream, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    cryptoStream.Write(data, 0, data.Length);
                }
            }
        }
    }
}