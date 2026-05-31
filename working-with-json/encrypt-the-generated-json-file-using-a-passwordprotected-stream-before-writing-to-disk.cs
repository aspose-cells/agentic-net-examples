using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

class Program
{
    static void Main()
    {
        // Sample object to be serialized to JSON
        var data = new { Name = "John Doe", Age = 30, Email = "john@example.com" };

        // Serialize the object to a JSON string
        string json = JsonSerializer.Serialize(data);
        byte[] jsonBytes = Encoding.UTF8.GetBytes(json);

        // Password used for encryption
        string password = "StrongPassword123!";

        // Generate a random salt (will be stored with the encrypted file)
        byte[] salt = new byte[16];
        using (var rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(salt);
        }

        // Derive a 256‑bit key and a 128‑bit IV from the password and salt
        using (var kdf = new Rfc2898DeriveBytes(password, salt, 100_000, HashAlgorithmName.SHA256))
        {
            byte[] key = kdf.GetBytes(32); // 256‑bit key
            byte[] iv = kdf.GetBytes(16);  // 128‑bit IV

            string outputPath = "encryptedData.json";

            // Write the salt followed by the encrypted JSON to the file
            using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                // Store salt first so it can be used during decryption
                fileStream.Write(salt, 0, salt.Length);

                using (Aes aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.IV = iv;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;

                    // Encrypt JSON bytes via CryptoStream
                    using (CryptoStream cryptoStream = new CryptoStream(fileStream, aes.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cryptoStream.Write(jsonBytes, 0, jsonBytes.Length);
                    }
                }
            }
        }

        Console.WriteLine("JSON file encrypted and saved successfully.");
    }
}