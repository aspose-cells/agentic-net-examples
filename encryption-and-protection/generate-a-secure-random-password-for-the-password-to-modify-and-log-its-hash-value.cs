// Title: C# code to generate a cryptographically secure random password and compute its SHA256 hash
// Common Searches: c# generate a cryptographically secure random password with RandomNumberGenerator | how to compute SHA256 hash of a string in .NET | log generated password and its SHA256 hash to console in C# | create a 16‑character strong password programmatically in C#
// Tags: RandomNumberGenerator password generation C# | SHA256 password hashing .NET | console logging of password hash C# | cryptographically strong password example C#

using System;
using System.Security.Cryptography;
using System.Text;

// The sample program creates a 16‑character password using a cryptographically secure random number generator, calculates its SHA256 hash, and writes both the password and the hash to the console.
class Program
{
    static void Main()
    {
        // Generate a secure random password
        string password = GenerateSecurePassword(16); // 16-character password

        // Compute SHA256 hash of the password
        string hash = ComputeSha256Hash(password);

        // Log the password and its hash
        Console.WriteLine($"Generated password: {password}");
        Console.WriteLine($"SHA256 hash: {hash}");
    }

    static string GenerateSecurePassword(int length)
    {
        const string validChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890!@#$%^&*()-_=+[]{}|;:,.<>?";
        var result = new StringBuilder();
        using (var rng = RandomNumberGenerator.Create())
        {
            byte[] uintBuffer = new byte[4];
            while (result.Length < length)
            {
                rng.GetBytes(uintBuffer);
                uint num = BitConverter.ToUInt32(uintBuffer, 0);
                result.Append(validChars[(int)(num % (uint)validChars.Length)]);
            }
        }
        return result.ToString();
    }

    static string ComputeSha256Hash(string rawData)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(rawData));
            var builder = new StringBuilder();
            foreach (byte b in bytes)
                builder.Append(b.ToString("x2"));
            return builder.ToString();
        }
    }
}
