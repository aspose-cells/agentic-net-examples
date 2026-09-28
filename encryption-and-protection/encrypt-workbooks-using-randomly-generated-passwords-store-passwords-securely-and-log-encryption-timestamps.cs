// Title: Encrypt an Excel workbook with a randomly generated password using Aspose.Cells for .NET, store the password with a timestamp, and log the encryption time
// AI Prompts: Write a C# method that generates a cryptographically strong random password, assigns it to Workbook.Settings.Password, and saves the workbook as a password‑protected XLSX file using Aspose.Cells. | Create a routine that appends the generated password together with the current UTC timestamp to a secure text file and records the encryption event with a timestamp in a separate log, handling any I/O exceptions gracefully.
// Common Searches: Aspose.Cells .NET generate random password for workbook protection | how to save a password‑protected Excel file and keep a password log in C# | store workbook passwords with timestamps using C# | log encryption timestamp for an Aspose.Cells workbook | C# encrypt Excel file with random password and record operation time
// Tags: Aspose.Cells workbook password generation | Workbook.Settings.Password usage | password archive with UTC timestamp | encryption event logging for Excel | C# cryptographic password generator for XLSX

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The example loads or creates an Excel workbook, generates a 12‑character cryptographically random password, applies it via Workbook.Settings.Password, saves the workbook as a protected XLSX file, appends the password with a UTC timestamp to a password archive, and writes an encryption timestamp to a log file, all with robust error handling for file operations.
class WorkbookEncryption
{
    static void Main()
    {
        // Paths for input workbook, encrypted output, password store, and log file
        string inputPath = "input.xlsx";
        string encryptedPath = "encrypted.xlsx";
        string passwordStorePath = "passwords.dat";
        string logPath = "encryption.log";

        try
        {
            // Ensure input workbook exists; create a simple one if missing
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath, SaveFormat.Xlsx);
            }

            // Load the workbook
            Workbook wb = new Workbook(inputPath);

            // Generate a random password
            string password = GenerateRandomPassword(12);

            // Apply password protection to the workbook
            wb.Settings.Password = password;
            wb.Save(encryptedPath, SaveFormat.Xlsx); // Save encrypted workbook

            // Store the password securely (plain text with timestamp for simplicity)
            StorePasswordSecurely(password, passwordStorePath);

            // Log the encryption timestamp
            LogEncryption(encryptedPath, logPath);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            try
            {
                File.AppendAllText(logPath, $"{DateTime.Now:O} - Error: {ex.Message}{Environment.NewLine}");
            }
            catch
            {
                // Suppress any logging failures
            }
        }
    }

    // Generates a random alphanumeric password of specified length
    static string GenerateRandomPassword(int length)
    {
        const string chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()";
        var sb = new StringBuilder();
        using (var rng = RandomNumberGenerator.Create())
        {
            byte[] buffer = new byte[4];
            while (sb.Length < length)
            {
                rng.GetBytes(buffer);
                uint num = BitConverter.ToUInt32(buffer, 0);
                sb.Append(chars[(int)(num % (uint)chars.Length)]);
            }
        }
        return sb.ToString();
    }

    // Stores the password with a UTC timestamp (plain text for demonstration)
    static void StorePasswordSecurely(string password, string storePath)
    {
        try
        {
            using (var fs = new FileStream(storePath, FileMode.Append, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(fs, Encoding.UTF8))
            {
                // Write UTC timestamp and password separated by a delimiter
                string entry = $"{DateTime.UtcNow:O}|{password}";
                writer.WriteLine(entry);
            }
        }
        catch (Exception ex)
        {
            // Optionally handle storage errors (e.g., log to console)
            Console.Error.WriteLine($"Failed to store password: {ex.Message}");
        }
    }

    // Appends an entry with the current timestamp to a log file
    static void LogEncryption(string workbookPath, string logPath)
    {
        try
        {
            string entry = $"{DateTime.Now:O} - Encrypted workbook: {workbookPath}";
            File.AppendAllText(logPath, entry + Environment.NewLine);
        }
        catch (Exception ex)
        {
            // Optionally handle logging errors
            Console.Error.WriteLine($"Failed to write log: {ex.Message}");
        }
    }
}
