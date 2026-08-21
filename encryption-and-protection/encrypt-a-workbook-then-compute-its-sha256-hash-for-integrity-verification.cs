// Title: Encrypt an Excel workbook with Aspose.Cells (256‑bit password) and compute its SHA‑256 checksum in C#
// Description: Creates a new Workbook, adds data, applies a strong 256‑bit password using Aspose.Cells encryption, saves the file, confirms encryption, and generates a SHA‑256 hash of the saved workbook for integrity verification.
// Keywords: Aspose.Cells encrypt workbook | 256‑bit Excel encryption C# | password protect Excel Aspose | SHA‑256 checksum Excel file | verify workbook integrity .NET | C# Excel file encryption | Aspose.Cells encryption options
// Common Searches: encrypt Excel file with Aspose.Cells C# | 256‑bit password protection for .xlsx using Aspose | calculate SHA‑256 hash of an Excel workbook in .NET | check if Aspose.Cells workbook is encrypted | Aspose.Cells generate checksum for encrypted file
// Developer Intent: Apply strong password protection to an Excel workbook with Aspose.Cells and obtain a SHA‑256 hash to confirm the file’s integrity.
// Use Cases: Secure confidential spreadsheets before uploading to a shared drive by encrypting them and storing a hash for later validation. | Create tamper‑evident financial reports: encrypt the workbook and keep the SHA‑256 checksum to detect any modifications. | Automate compliance checks by comparing the computed hash of a downloaded encrypted Excel file against a known value.
// AI Prompts: Generate C# code that encrypts an Aspose.Cells workbook with a custom password using 256‑bit encryption and returns the SHA‑256 hash of the saved file. | Show how to open an existing encrypted Excel workbook with Aspose.Cells, verify the password, and recompute its SHA‑256 checksum for integrity checking. | Explain the steps to combine Aspose.Cells encryption settings with .NET cryptography to produce a hash for a protected workbook.

using System;
using System.IO;
using System.Security.Cryptography;
using Aspose.Cells;

// Creates a new Workbook, adds data, applies a strong 256‑bit password using Aspose.Cells encryption, saves the file, confirms encryption, and generates a SHA‑256 hash of the saved workbook for integrity verification.
public class EncryptAndHashDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and ensure it is disposed properly
        using (Workbook workbook = new Workbook())
        {
            // Add sample data
            workbook.Worksheets[0].Cells["A1"].PutValue("Sensitive data");

            // Set a password to encrypt the workbook
            workbook.Settings.Password = "StrongPassword123";

            // Specify encryption options (strong encryption with 256‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the encrypted workbook to disk
            string filePath = "EncryptedWorkbook.xlsx";
            workbook.Save(filePath);

            // Verify that the workbook is encrypted
            Console.WriteLine($"Is workbook encrypted: {workbook.Settings.IsEncrypted}");

            // Compute SHA‑256 hash of the saved file for integrity verification
            if (File.Exists(filePath))
            {
                byte[] fileBytes = File.ReadAllBytes(filePath);
                using (SHA256 sha256 = SHA256.Create())
                {
                    byte[] hashBytes = sha256.ComputeHash(fileBytes);
                    string hashString = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                    Console.WriteLine($"SHA‑256 hash: {hashString}");
                }
            }
            else
            {
                Console.WriteLine($"File not found: {filePath}");
            }
        }
    }
}
