// Title: C# – Encrypt an Aspose.Cells workbook with a random password, log the timestamp, and store the password securely
// Description: Demonstrates how to create a workbook, generate a 16‑character random alphanumeric password, apply StrongCryptographicProvider (128‑bit) encryption, save the file, write a UTC timestamp to a log, and encrypt the password with AES (SHA‑256 derived key) before saving it to a binary store.
// Keywords: Aspose.Cells encrypt workbook C# | random password generation .NET | StrongCryptographicProvider 128‑bit | AES password storage | SHA‑256 key derivation | encryption timestamp log | secure Excel file protection | C# workbook security example
// Common Searches: how to encrypt an Excel file with Aspose.Cells using a random password | C# log workbook encryption timestamp | store Aspose.Cells password securely with AES | set encryption type and key size for Aspose.Cells workbook | generate random password for Excel encryption .NET
// Developer Intent: Protect a workbook with a unique password, record the exact encryption time, and keep the password safely encrypted for later retrieval.
// Use Cases: Compliance‑driven financial reports that require per‑file passwords and audit‑ready timestamps. | Automated nightly backups of sensitive spreadsheets, each encrypted with a distinct password stored in an encrypted vault. | Web services that deliver password‑protected Excel files while managing passwords on the server side.
// AI Prompts: Generate C# code to decrypt the AES‑encrypted password file and open the protected workbook with Aspose.Cells. | Refactor StorePasswordSecurely to use a random salt and PBKDF2‑derived key instead of a static fallback key. | Create a unit test that confirms the workbook is saved with encryption enabled and that the log entry follows the ISO 8601 UTC format.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// Demonstrates how to create a workbook, generate a 16‑character random alphanumeric password, apply StrongCryptographicProvider (128‑bit) encryption, save the file, write a UTC timestamp to a log, and encrypt the password with AES (SHA‑256 derived key) before saving it to a binary store.
class WorkbookEncryptionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Sensitive Data");

            // Generate a random password (16 characters)
            string password = GenerateRandomPassword(16);

            // Apply password and encryption options to the workbook
            wb.Settings.Password = password;
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Define the output file path
            string filePath = "EncryptedWorkbook.xlsx";

            // Save the encrypted workbook
            wb.Save(filePath);

            // Log the encryption timestamp
            string logEntry = $"{DateTime.UtcNow:o} - Workbook encrypted and saved to {filePath}";
            Console.WriteLine(logEntry);
            File.AppendAllText("encryption_log.txt", logEntry + Environment.NewLine);

            // Securely store the password
            StorePasswordSecurely(password, "password_store.bin");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Generates a random alphanumeric password of the specified length
    static string GenerateRandomPassword(int length)
    {
        const string chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
        {
            byte[] data = new byte[length];
            rng.GetBytes(data);
            char[] result = new char[length];
            for (int i = 0; i < length; i++)
            {
                result[i] = chars[data[i] % chars.Length];
            }
            return new string(result);
        }
    }

    // Encrypts the password and writes it to a file using AES.
    static void StorePasswordSecurely(string password, string filePath)
    {
        try
        {
            byte[] passwordBytes = Encoding.UTF8.GetBytes(password);
            byte[] encrypted;

            // AES encryption (cross‑platform)
            using (Aes aes = Aes.Create())
            {
                // Derive a key from a static passphrase (for demo purposes only)
                aes.Key = SHA256.Create().ComputeHash(Encoding.UTF8.GetBytes("fallback-key"));
                aes.IV = new byte[16]; // Zero IV (not ideal for production)

                using (ICryptoTransform encryptor = aes.CreateEncryptor())
                {
                    encrypted = encryptor.TransformFinalBlock(passwordBytes, 0, passwordBytes.Length);
                }
            }

            File.WriteAllBytes(filePath, encrypted);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to store password securely: {ex.Message}");
        }
    }
}
