// Title: Encrypt an Excel workbook with a random password using Aspose.Cells for .NET, securely store the password, and log the encryption timestamp
// Common Searches: C# Aspose.Cells encrypt workbook with random password and save as Xlsx | How to generate a strong random password for Excel file protection using Aspose.Cells .NET | Store generated workbook passwords securely in a file with Aspose.Cells example | Log encryption timestamp and password hash when protecting an Excel workbook in C# | Aspose.Cells encrypt Excel workbook and maintain password log for audit
// Tags: Aspose.Cells random password workbook encryption | secure password storage for encrypted Excel files .NET | encryption timestamp logging Aspose.Cells | cryptographically strong password generation C# | save encrypted workbook as Xlsx using Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace WorkbookEncryptionDemo
{
    // Helper class for password generation, secure storage and logging
    // The sample loads an Excel file with Aspose.Cells, generates a cryptographically random password, applies it via Workbook.Settings.Password, saves the workbook as an encrypted Xlsx, stores the password in a simple encoded file, and appends a log entry with the UTC timestamp and SHA‑256 hash of the password.
    public static class EncryptionHelper
    {
        // Path to the file that stores encrypted passwords (simple key‑value store)
        private static readonly string PasswordStorePath = "encrypted_passwords.dat";

        // Path to the encryption log file
        private static readonly string LogFilePath = "encryption_log.txt";

        // In‑memory cache of file‑to‑encrypted‑password mappings
        private static readonly Dictionary<string, byte[]> PasswordCache = LoadPasswordCache();

        // Generate a random password of the specified length (default 12 characters)
        public static string GenerateRandomPassword(int length = 12)
        {
            const string validChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()";
            var password = new StringBuilder();
            using (var rng = RandomNumberGenerator.Create())
            {
                byte[] uintBuffer = new byte[4];
                while (password.Length < length)
                {
                    rng.GetBytes(uintBuffer);
                    uint num = BitConverter.ToUInt32(uintBuffer, 0);
                    password.Append(validChars[(int)(num % (uint)validChars.Length)]);
                }
            }
            return password.ToString();
        }

        // Store the password for a given workbook file name (simple reversible encoding for demo)
        public static void StorePassword(string workbookFileName, string password)
        {
            // Encode password bytes (replace with real DPAPI or other secure method in production)
            byte[] encoded = Encoding.UTF8.GetBytes(password);
            PasswordCache[workbookFileName] = encoded;
            SavePasswordCache();
        }

        // Retrieve a stored password (if needed)
        public static string RetrievePassword(string workbookFileName)
        {
            if (PasswordCache.TryGetValue(workbookFileName, out byte[] encoded))
            {
                return Encoding.UTF8.GetString(encoded);
            }
            return null;
        }

        // Log the encryption event with timestamp
        public static void LogEncryption(string workbookFileName, string password)
        {
            string logEntry = $"{DateTime.UtcNow:O}\t{workbookFileName}\tPasswordHash:{ComputeHash(password)}";
            try
            {
                File.AppendAllText(LogFilePath, logEntry + Environment.NewLine);
            }
            catch
            {
                // Swallow logging errors to avoid breaking the main flow
            }
        }

        // Compute a SHA‑256 hash of the password for logging (do not store plain text)
        private static string ComputeHash(string input)
        {
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(input));
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }

        // Load existing encrypted passwords from disk (if file exists)
        private static Dictionary<string, byte[]> LoadPasswordCache()
        {
            var dict = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            if (File.Exists(PasswordStorePath))
            {
                foreach (var line in File.ReadAllLines(PasswordStorePath))
                {
                    // Each line format: filename|base64(encodedPassword)
                    var parts = line.Split('|');
                    if (parts.Length == 2)
                    {
                        string fileName = parts[0];
                        byte[] encoded = Convert.FromBase64String(parts[1]);
                        dict[fileName] = encoded;
                    }
                }
            }
            return dict;
        }

        // Persist the in‑memory password cache to disk
        private static void SavePasswordCache()
        {
            var sb = new StringBuilder();
            foreach (var kvp in PasswordCache)
            {
                sb.AppendLine($"{kvp.Key}|{Convert.ToBase64String(kvp.Value)}");
            }
            try
            {
                File.WriteAllText(PasswordStorePath, sb.ToString());
            }
            catch
            {
                // Swallow persistence errors; in production handle appropriately
            }
        }
    }

    public class WorkbookEncryptor
    {
        // Encrypts the workbook at inputPath, saves to outputPath, stores password securely and logs the operation
        public static void EncryptWorkbook(string inputPath, string outputPath)
        {
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            try
            {
                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Generate a random password
                string password = EncryptionHelper.GenerateRandomPassword();

                // Apply password protection using WorkbookSettings
                workbook.Settings.Password = password;

                // Save the encrypted workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                // Store the password securely
                EncryptionHelper.StorePassword(Path.GetFileName(outputPath), password);

                // Log the encryption timestamp and password hash
                EncryptionHelper.LogEncryption(Path.GetFileName(outputPath), password);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error encrypting workbook: {ex.Message}");
                throw;
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main(string[] args)
        {
            // Paths can be adjusted as needed
            string sourceFile = "Sample.xlsx";
            string encryptedFile = "Sample_Encrypted.xlsx";

            try
            {
                WorkbookEncryptor.EncryptWorkbook(sourceFile, encryptedFile);
                Console.WriteLine("Workbook encrypted successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Operation failed: {ex.Message}");
            }
        }
    }
}
