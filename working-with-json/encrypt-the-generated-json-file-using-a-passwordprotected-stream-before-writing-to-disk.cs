using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

class JsonEncryptionExample
{
    // Password used to protect the JSON file
    private const string Password = "StrongPassword123!";

    // Number of iterations for the key derivation function
    private const int DerivationIterations = 100_000;

    // Size of the random salt (in bytes)
    private const int SaltSize = 16;

    // Size of the initialization vector (in bytes) for AES
    private const int IvSize = 16;

    static void Main()
    {
        // Sample JSON content to be encrypted
        string jsonContent = @"{
    ""Name"": ""John Doe"",
    ""Age"": 30,
    ""Email"": ""john.doe@example.com"",
    ""Roles"": [""Admin"", ""User""]
}";

        // Convert JSON string to bytes
        byte[] plainBytes = Encoding.UTF8.GetBytes(jsonContent);

        // Encrypt the JSON bytes using a password‑derived AES key
        byte[] encryptedData = EncryptWithPassword(plainBytes, Password);

        // Write the encrypted data to disk (the file contains: [salt][IV][ciphertext])
        string outputPath = "encryptedData.bin";
        File.WriteAllBytes(outputPath, encryptedData);

        Console.WriteLine($"Encrypted JSON written to '{outputPath}'.");
    }

    /// <summary>
    /// Encrypts the supplied data using AES (CBC) with a key derived from the given password.
    /// The returned byte array layout is: [salt][IV][ciphertext].
    /// </summary>
    private static byte[] EncryptWithPassword(byte[] data, string password)
    {
        // Generate a random salt
        byte[] salt = new byte[SaltSize];
        using (var rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(salt);
        }

        // Derive a 256‑bit key from the password and salt
        using var keyDerivation = new Rfc2898DeriveBytes(password, salt, DerivationIterations, HashAlgorithmName.SHA256);
        byte[] key = keyDerivation.GetBytes(32); // 256 bits

        // Create AES algorithm instance
        using var aes = Aes.Create();
        aes.KeySize = 256;
        aes.BlockSize = 128;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.PKCS7;
        aes.Key = key;

        // Generate a random IV
        aes.GenerateIV();
        byte[] iv = aes.IV;

        // Perform encryption
        using var encryptor = aes.CreateEncryptor();
        byte[] cipherText;
        using (var ms = new MemoryStream())
        {
            using (var cryptoStream = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
            {
                cryptoStream.Write(data, 0, data.Length);
                cryptoStream.FlushFinalBlock();
                cipherText = ms.ToArray();
            }
        }

        // Combine salt + IV + ciphertext into a single byte array
        byte[] result = new byte[SaltSize + IvSize + cipherText.Length];
        Buffer.BlockCopy(salt, 0, result, 0, SaltSize);
        Buffer.BlockCopy(iv, 0, result, SaltSize, IvSize);
        Buffer.BlockCopy(cipherText, 0, result, SaltSize + IvSize, cipherText.Length);
        return result;
    }
}