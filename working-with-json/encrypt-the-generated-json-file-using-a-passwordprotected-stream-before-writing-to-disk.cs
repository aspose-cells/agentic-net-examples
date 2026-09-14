// Title: Encrypt JSON output from an Aspose.Cells workbook using a password‑protected AES stream in C#
// AI Prompts: Create a method that accepts a Stream and a password, derives a 256‑bit AES key with Rfc2898DeriveBytes, encrypts the stream in CBC mode, and returns the encrypted byte array. | Update the Aspose.Cells example to write the encrypted byte array to a file after securing the JSON data with a password‑derived key. | Add decryption logic that reads the encrypted file, extracts the prefixed salt, derives the same key, and restores the original JSON content.
// Common Searches: how to encrypt Aspose.Cells JSON output with AES in C# | C# encrypt memory stream using password and Rfc2898DeriveBytes | store encrypted JSON file from Aspose.Cells workbook .NET Core
// Tags: Aspose.Cells JSON secure export C# | AES CBC stream encryption .NET | Rfc2898DeriveBytes key derivation C# | password‑protected JSON file Aspose.Cells | memory stream encryption Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Security.Cryptography;

// The sample creates a workbook, saves it as JSON to a memory stream, encrypts the JSON using AES with a password‑derived key (salt prefixed), and writes the encrypted bytes to 'output_encrypted.json'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a workbook and populate it with sample data
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Age");
            cells["A2"].PutValue("Alice");
            cells["B2"].PutValue(30);
            cells["A3"].PutValue("Bob");
            cells["B3"].PutValue(25);

            // Save the workbook as JSON into a memory stream
            using (var jsonStream = new MemoryStream())
            {
                workbook.Save(jsonStream, SaveFormat.Json);
                jsonStream.Position = 0; // Reset stream position for reading

                // Encrypt the JSON data with a password‑protected stream
                const string password = "MySecretPassword";
                byte[] encryptedBytes = EncryptStream(jsonStream, password);

                // Write the encrypted bytes to disk
                File.WriteAllBytes("output_encrypted.json", encryptedBytes);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Encrypts the input stream using AES with a key derived from the password.
    // The generated salt is prefixed to the output so it can be used for decryption.
    static byte[] EncryptStream(Stream input, string password)
    {
        // Generate a random salt
        byte[] salt = new byte[16];
        using (var rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(salt);
        }

        // Derive a 256‑bit key and a 128‑bit IV from the password and salt
        using var kdf = new Rfc2898DeriveBytes(password, salt, 100_000, HashAlgorithmName.SHA256);
        byte[] key = kdf.GetBytes(32); // 256‑bit key
        byte[] iv = kdf.GetBytes(16);  // 128‑bit IV

        using (var aes = Aes.Create())
        {
            aes.Key = key;
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;

            using (var output = new MemoryStream())
            {
                // Write the salt first so it can be read during decryption
                output.Write(salt, 0, salt.Length);

                // Encrypt the JSON data
                using (var cryptoStream = new CryptoStream(output, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    input.CopyTo(cryptoStream);
                }

                return output.ToArray();
            }
        }
    }
}
