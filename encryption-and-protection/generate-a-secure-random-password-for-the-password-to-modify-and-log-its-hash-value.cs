// Title: Create a 16‑character cryptographically secure password and log its SHA‑256 hash in C#
// AI Prompts: Generate a C# method that returns a random password of a specified length using RandomNumberGenerator and a custom character set. | Write a C# function that accepts a string and returns its SHA256 hash as a hexadecimal string. | Update the sample to output the SHA256 hash in Base64 instead of hex and suppress the plain‑text password in the console.
// Common Searches: c# generate random password with cryptographic RNG and custom symbols | how to compute SHA256 hash of a string in .NET Core and get hex output | best practice for logging password hashes without exposing the password in C# applications
// Tags: cryptographic random password generation C# | SHA256 hash computation .NET | RandomNumberGenerator password example | hexadecimal hash output C# | Base64 hash representation .NET

using System;
using System.Security.Cryptography;
using System.Text;

// The example creates a 16‑character password using a cryptographically secure RandomNumberGenerator, computes its SHA‑256 hash as a hex string, and writes the hash (and optionally the password) to the console.
class Program
{
    static void Main()
    {
        // Generate a secure random password
        string password = GenerateSecurePassword(16);

        // Compute SHA256 hash of the password
        string hash = ComputeSha256Hash(password);

        // Log the hash value
        Console.WriteLine($"Password Hash (SHA256): {hash}");

        // Optional: display the generated password (remove in production)
        Console.WriteLine($"Generated Password: {password}");
    }

    // Generates a random password of the specified length using a cryptographically secure RNG
    static string GenerateSecurePassword(int length)
    {
        const string validChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!@#$%^&*()-_=+[]{}|;:,.<>?";
        char[] password = new char[length];
        byte[] uintBuffer = new byte[4];

        using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
        {
            for (int i = 0; i < length; i++)
            {
                rng.GetBytes(uintBuffer);
                uint num = BitConverter.ToUInt32(uintBuffer, 0);
                password[i] = validChars[(int)(num % (uint)validChars.Length)];
            }
        }

        return new string(password);
    }

    // Computes the SHA256 hash of the input string and returns it as a hex string
    static string ComputeSha256Hash(string rawData)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(rawData));
            StringBuilder builder = new StringBuilder();
            foreach (byte b in bytes)
                builder.Append(b.ToString("x2"));
            return builder.ToString();
        }
    }
}
