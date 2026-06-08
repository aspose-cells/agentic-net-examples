using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

namespace WorkbookEncryptionDemo
{
    class Program
    {
        // In‑memory store for passwords (could be replaced with a secure vault)
        private static readonly Dictionary<string, (string Password, DateTime Timestamp)> PasswordStore
            = new Dictionary<string, (string, DateTime)>();

        static void Main()
        {
            // 1. Create a new workbook and add sample data
            Workbook workbook = new Workbook();                         // create
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // 2. Generate a random password
            string randomPassword = GenerateRandomPassword(16);        // 16‑byte password

            // 3. Apply encryption settings
            workbook.Settings.Password = randomPassword;               // set workbook password
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // 4. Save the encrypted workbook
            string filePath = "EncryptedWorkbook.xlsx";
            workbook.Save(filePath);                                    // save

            // 5. Store the password securely with a timestamp
            StorePassword(filePath, randomPassword);

            // 6. Log encryption timestamp
            Console.WriteLine($"Workbook encrypted at {DateTime.UtcNow:u}");
            Console.WriteLine($"Password stored for file: {filePath}");

            // 7. Demonstrate loading the encrypted workbook
            LoadOptions loadOptions = new LoadOptions { Password = randomPassword };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions); // load
            Console.WriteLine($"Loaded cell value: {loadedWorkbook.Worksheets[0].Cells["A1"].Value}");
        }

        // Generates a random password of the specified byte length and returns it as a Base64 string
        private static string GenerateRandomPassword(int byteLength)
        {
            byte[] bytes = new byte[byteLength];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }
            return Convert.ToBase64String(bytes);
        }

        // Stores the password and timestamp in the in‑memory dictionary and optionally persists to a file
        private static void StorePassword(string filePath, string password)
        {
            DateTime now = DateTime.UtcNow;
            PasswordStore[filePath] = (password, now);

            // Example of persisting to a simple text file (not encrypted for demo purposes)
            string logLine = $"{filePath}|{password}|{now:o}";
            File.AppendAllLines("PasswordLog.txt", new[] { logLine });
        }
    }
}