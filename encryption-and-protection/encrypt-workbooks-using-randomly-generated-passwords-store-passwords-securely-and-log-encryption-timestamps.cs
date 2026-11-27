// Title: Encrypt an Excel workbook with a random password using Aspose.Cells for .NET, save the password to a file, and record the encryption timestamp
// Common Searches: Aspose.Cells encrypt Excel file with random password C# example | How to save generated workbook password to a file in .NET | Logging encryption time for password‑protected Excel workbook in C# | Generate strong random password for Excel workbook protection using Aspose.Cells | Store Excel workbook password securely after encryption with Aspose.Cells
// Tags: Aspose.Cells workbook password protection | random password generation C# | persist workbook password .NET | encryption timestamp logging C# | Excel file encryption Aspose.Cells | secure password storage .NET example

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The sample generates a 16‑character strong random password, loads an existing .xlsx workbook with Aspose.Cells, applies the password via Workbook.Settings.Password, saves the encrypted workbook to a new location, writes the password to a file (demonstration purpose), and appends a UTC timestamp entry to a log file indicating when the encryption occurred.
class WorkbookEncryption
{
    // Path to the original workbook
    private const string InputFilePath = @"C:\Data\OriginalWorkbook.xlsx";

    // Path where the encrypted workbook will be saved
    private const string EncryptedFilePath = @"C:\Data\EncryptedWorkbook.xlsx";

    // Path to store the protected password (binary)
    private const string PasswordStorePath = @"C:\Data\WorkbookPassword.bin";

    // Path to the encryption log file
    private const string LogFilePath = @"C:\Data\EncryptionLog.txt";

    static void Main()
    {
        try
        {
            // Generate a strong random password
            string password = GenerateRandomPassword(16);

            // Load the workbook, apply password protection, and save it
            EncryptWorkbook(InputFilePath, EncryptedFilePath, password);

            // Store the password securely (simple storage for demo purposes)
            StorePasswordSecurely(password, PasswordStorePath);

            // Log the encryption timestamp
            LogEncryptionTimestamp(EncryptedFilePath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Generates a random alphanumeric password of the specified length
    private static string GenerateRandomPassword(int length)
    {
        const string validChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+";
        char[] password = new char[length];
        byte[] randomBytes = new byte[length];

        using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
        {
            rng.GetBytes(randomBytes);
        }

        for (int i = 0; i < length; i++)
        {
            int index = randomBytes[i] % validChars.Length;
            password[i] = validChars[index];
        }

        return new string(password);
    }

    // Loads a workbook, sets its password, and saves it to a new file
    private static void EncryptWorkbook(string sourcePath, string destPath, string password)
    {
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("Source workbook not found.", sourcePath);

        // Ensure destination directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(destPath));

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Apply password protection
        workbook.Settings.Password = password;

        // Save the encrypted workbook
        workbook.Save(destPath);
    }

    // Stores the password securely (for demonstration, writes plain text; replace with DPAPI or other secure method as needed)
    private static void StorePasswordSecurely(string password, string storePath)
    {
        // Ensure directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(storePath));

        // Write the password to the file (replace with proper encryption in production)
        File.WriteAllText(storePath, password, Encoding.UTF8);
    }

    // Appends a timestamp entry to the log file indicating when encryption occurred
    private static void LogEncryptionTimestamp(string encryptedFilePath)
    {
        // Ensure log directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));

        string logEntry = $"{DateTime.UtcNow:O} - Encrypted workbook saved to '{encryptedFilePath}'";
        File.AppendAllLines(LogFilePath, new[] { logEntry });
    }
}
