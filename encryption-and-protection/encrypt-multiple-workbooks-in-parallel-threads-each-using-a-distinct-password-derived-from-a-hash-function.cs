// Title: Encrypt multiple Excel workbooks in parallel with Aspose.Cells using a SHA‑256 derived password per file (C#)
// AI Prompts: Write C# code that reads a list of .xlsx file paths, creates a unique password for each by hashing the file name with a secret key using SHA‑256, and encrypts the workbooks concurrently with Aspose.Cells. | Enhance the parallel encryption sample to write each generated password to an encrypted log file after the workbook is saved. | Modify the password‑derivation function to include the file's creation timestamp together with the secret key, then encrypt the workbooks in parallel.
// Common Searches: parallel encryption of multiple Excel files with Aspose.Cells in C# | generate per‑file SHA256 password for Excel workbook using Aspose.Cells | set opening password for Aspose.Cells workbook programmatically in .NET | how to save encrypted Excel workbook as Xlsx using Aspose.Cells
// Tags: parallel workbook encryption Aspose.Cells | SHA256 password derivation for Excel | Aspose.Cells opening password protection C# | save encrypted workbook as Xlsx Aspose.Cells | multithreaded Excel file protection .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

// The example loads each .xlsx file from a list, derives a unique password by hashing the file name with a secret key using SHA‑256, applies the password via Workbook.Settings.Password, and saves the protected copy as a new Xlsx file. All files are processed concurrently with Parallel.ForEach for fast, thread‑safe encryption.
class Program
{
    static void Main()
    {
        // List of workbook file paths to encrypt
        var files = new List<string>
        {
            "Book1.xlsx",
            "Book2.xlsx",
            // Add additional workbook paths as needed
        };

        // Secret key used in password derivation (keep it safe)
        const string secret = "MySecretKey";

        // Encrypt each workbook in parallel
        Parallel.ForEach(files, filePath =>
        {
            try
            {
                // Verify that the source file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Derive a unique password for this workbook
                string password = DerivePassword(filePath, secret);

                // Load the workbook
                var workbook = new Workbook(filePath);

                // Apply opening password protection
                workbook.Settings.Password = password;

                // Define output path (you can overwrite the original if desired)
                string directory = Path.GetDirectoryName(filePath) ?? Directory.GetCurrentDirectory();
                string encryptedPath = Path.Combine(
                    directory,
                    Path.GetFileNameWithoutExtension(filePath) + "_encrypted.xlsx");

                // Save the encrypted workbook (no special save options needed)
                workbook.Save(encryptedPath, SaveFormat.Xlsx);
                Console.WriteLine($"Encrypted file saved: {encryptedPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {filePath}: {ex.Message}");
            }
        });
    }

    // Generates a password by hashing the file name with a secret
    static string DerivePassword(string filePath, string secret)
    {
        using (SHA256 sha = SHA256.Create())
        {
            // Combine file name and secret to create input for hashing
            string input = Path.GetFileName(filePath) + secret;
            byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(input));

            // Use a portion of the hash and encode as Base64 for a readable password
            return Convert.ToBase64String(hash, 0, 16);
        }
    }
}
