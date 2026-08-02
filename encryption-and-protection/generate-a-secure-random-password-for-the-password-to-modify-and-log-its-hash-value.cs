// Title: Aspose.Cells .NET: Generate Secure Random Password, Apply Write & Worksheet Protection, Retrieve Hash
// Description: Demonstrates how to create a new Workbook, generate a cryptographically‑secure 16‑byte Base64 password with RandomNumberGenerator, set it as the file's write‑protection and a worksheet's protection, obtain the worksheet password hash via GetPasswordHash, log both values, and save the workbook.
// Keywords: Aspose.Cells C# example | generate secure random password .NET | Excel write protection Aspose.Cells | worksheet protection password hash | RandomNumberGenerator Fill | GetPasswordHash method | Base64 password Excel | cryptographic password Excel file
// Common Searches: how to set a random write‑protection password in Aspose.Cells | retrieve worksheet protection password hash C# Aspose.Cells | generate cryptographically secure password for Excel workbook .NET | Aspose.Cells protect worksheet with password and get hash | C# example for Excel file write protection using Aspose
// Developer Intent: Create a workbook protected by a strong random password and obtain the worksheet protection hash for verification or logging.
// Use Cases: Generate a 16‑byte secure password, apply it to workbook write‑protection and a worksheet, then save the protected file. | Protect all cells on a worksheet with the same random password and retrieve its integer hash for audit purposes. | Log the Base64 password and its hash before saving to support compliance and troubleshooting.
// AI Prompts: Show C# code that creates a 16‑byte secure password, sets it as write‑protection and worksheet protection in Aspose.Cells, and prints the worksheet password hash. | Explain how Aspose.Cells Protection.GetPasswordHash calculates the hash and how to compare it with a known password. | Provide a step‑by‑step guide to protect an Excel workbook with a random password using Aspose.Cells and retrieve the hash for validation.

using System;
using System.Security.Cryptography;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a new Workbook, generate a cryptographically‑secure 16‑byte Base64 password with RandomNumberGenerator, set it as the file's write‑protection and a worksheet's protection, obtain the worksheet password hash via GetPasswordHash, log both values, and save the workbook.
    public class SecureRandomPasswordDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Generate a secure random password (16 bytes, Base64 encoded)
                byte[] passwordBytes = new byte[16];
                RandomNumberGenerator.Fill(passwordBytes);
                string randomPassword = Convert.ToBase64String(passwordBytes);

                // Set the write‑protection password (required to modify the file)
                workbook.Settings.WriteProtection.Password = randomPassword;

                // Optionally protect the first worksheet with the same password
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Protection.Password = randomPassword;
                worksheet.Protect(ProtectionType.All);

                // Retrieve the hash of the worksheet protection password
                int passwordHash = worksheet.Protection.GetPasswordHash();

                // Log the generated password and its hash
                Console.WriteLine($"Generated password: {randomPassword}");
                Console.WriteLine($"Password hash (worksheet protection): {passwordHash}");

                // Save the workbook
                workbook.Save("SecureRandomPasswordDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SecureRandomPasswordDemo.Run();
        }
    }
}
