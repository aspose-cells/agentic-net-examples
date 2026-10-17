// Title: Create a 16‑character cryptographically secure password and display its SHA‑256 hash using C# and .NET
// Common Searches: C# generate random password with RandomNumberGenerator and compute SHA256 hash | how to hash a generated password using SHA256 in .NET Core | secure password generation example using RNGCryptoServiceProvider in C# | log password hash to console in a C# application | C# create strong password and output its hexadecimal hash
// Tags: RandomNumberGenerator secure password generation C# | SHA256 hash computation .NET | hexadecimal hash output console C# | cryptographic random password example | C# password hash logging

using System;
using System.Security.Cryptography;
using System.Text;

// The sample program generates a 16‑character cryptographically secure password with RandomNumberGenerator, computes its SHA‑256 hash using SHA256.Create, and writes the hexadecimal hash to the console.
class Program
{
    static void Main()
    {
        // Generate a secure random password
        string password = GenerateSecurePassword(16); // 16-character password

        // Compute SHA256 hash of the password
        string hash = ComputeSha256Hash(password);

        // Log the hash value
        Console.WriteLine($"Password hash (SHA256): {hash}");
    }

    // Generates a random password of the specified length using a cryptographically secure RNG
    static string GenerateSecurePassword(int length)
    {
        const string validChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+[]{}|;:,.<>?";
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
