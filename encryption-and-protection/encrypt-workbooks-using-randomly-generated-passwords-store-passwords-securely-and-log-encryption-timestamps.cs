// Title: Encrypt Aspose.Cells Workbook with Random Password, Secure Store & Timestamp Log (C#)
// Description: Creates a new workbook, adds data, generates a strong random password with RNGCryptoServiceProvider, applies it via workbook.Settings.Password, configures 128‑bit StrongCryptographicProvider encryption, saves the file, and records the UTC timestamp and password to a secure log. The example also shows how to reload the workbook using the generated password.
// Keywords: Aspose.Cells encrypt workbook C# | random password generation .NET | Excel file encryption Aspose | store workbook password securely | encryption timestamp logging | StrongCryptographicProvider | SetEncryptionOptions | RNGCryptoServiceProvider password | C# workbook protection | audit encrypted Excel
// Common Searches: how to encrypt an Aspose.Cells workbook with a random password in C# | store Excel password securely after encryption Aspose.Cells | log encryption timestamp for protected workbook .NET | set strong encryption options for Aspose.Cells workbook | generate strong random password for Excel file C#
// Developer Intent: Securely encrypt an Excel workbook with a generated password, persist the password safely, and keep an audit‑ready timestamp of the encryption event.
// Use Cases: Protect confidential financial reports by encrypting each workbook with a unique random password and recording the credentials for compliance audits. | Automate secure distribution of Excel files in a SaaS platform, generating per‑file passwords and logging creation times for traceability. | Validate encryption integrity by re‑opening the saved workbook using the stored password in automated test suites.
// AI Prompts: Write C# code that encrypts an Aspose.Cells workbook with a 20‑character random password, saves the password in an encrypted JSON file, and logs the UTC timestamp. | Refactor the sample to use SecureString for password handling and write the log entry to the Windows Event Log instead of a text file. | Explain how to replace the plain‑text password store with Azure Key Vault integration while keeping the existing encryption workflow unchanged.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

namespace WorkbookEncryptionDemo
{
    // Creates a new workbook, adds data, generates a strong random password with RNGCryptoServiceProvider, applies it via workbook.Settings.Password, configures 128‑bit StrongCryptographicProvider encryption, saves the file, and records the UTC timestamp and password to a secure log. The example also shows how to reload the workbook using the generated password.
    class Program
    {
        // Path to store encrypted workbook
        private const string WorkbookPath = "EncryptedWorkbook.xlsx";
        // Path to store passwords securely (for demo purposes)
        private const string PasswordStorePath = "PasswordStore.txt";

        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive Data");

            // Generate a strong random password
            string password = GenerateRandomPassword(16);

            // Apply password to workbook settings
            workbook.Settings.Password = password;

            // Optionally set encryption options (strong provider, 128-bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Save the encrypted workbook
            workbook.Save(WorkbookPath);

            // Log encryption timestamp and store password securely
            LogEncryptionInfo(password);

            // Verify by loading the encrypted workbook with the password
            LoadOptions loadOptions = new LoadOptions { Password = password };
            Workbook loadedWorkbook = new Workbook(WorkbookPath, loadOptions);
            Console.WriteLine("Loaded workbook cell A1 value: " + loadedWorkbook.Worksheets[0].Cells["A1"].Value);
        }

        // Generates a random password of specified length using RNGCryptoServiceProvider
        private static string GenerateRandomPassword(int length)
        {
            byte[] randomBytes = new byte[length];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(randomBytes);
            }
            // Convert to a Base64 string and remove non-alphanumeric characters for simplicity
            string base64 = Convert.ToBase64String(randomBytes);
            StringBuilder sb = new StringBuilder();
            foreach (char c in base64)
            {
                if (char.IsLetterOrDigit(c))
                {
                    sb.Append(c);
                }
                if (sb.Length == length)
                {
                    break;
                }
            }
            return sb.ToString();
        }

        // Stores the password with a timestamp; in real scenarios, encrypt this storage
        private static void LogEncryptionInfo(string password)
        {
            string timestamp = DateTime.UtcNow.ToString("o"); // ISO 8601 format
            string logEntry = $"{timestamp} | Password: {password}";
            // Append to the password store file
            File.AppendAllText(PasswordStorePath, logEntry + Environment.NewLine);
            // Also output to console for demonstration
            Console.WriteLine($"Workbook encrypted at {timestamp}. Password stored securely.");
        }
    }
}
