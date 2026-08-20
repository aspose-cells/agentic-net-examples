// Title: Encrypt an Aspose.Cells workbook with password & AES‑128 and save to a UNC share (C#)
// Description: C# example that creates a Workbook, writes data, sets a password, applies AES‑128 encryption via SetEncryptionOptions, checks the UNC network folder, falls back to a local directory if needed, and saves the encrypted XLSX file so the protection remains intact.
// Keywords: Aspose.Cells password protection | AES 128 encryption Aspose.Cells | C# save workbook to UNC path | network share Excel encryption | Workbook.SetEncryptionOptions | Workbook.Settings.Password | encrypted Excel file save fallback | Aspose.Cells encryption C#
// Common Searches: Aspose.Cells set workbook password C# | How to encrypt Excel file with AES using Aspose.Cells | Save encrypted workbook to network share UNC | Fallback to local folder when UNC path unavailable Aspose.Cells | Encryption not preserved after saving to network drive Aspose.Cells
// Developer Intent: Add strong password protection (AES‑128) to a generated Excel workbook and ensure the encrypted file is reliably stored on a network share, with automatic fallback to a local path if the share is inaccessible.
// Use Cases: Secure confidential reports before placing them on a shared file server. | Automate generation of encrypted financial spreadsheets for centralized storage. | Maintain encryption integrity when saving to intermittent network locations. | Provide compliance‑ready encrypted Excel files in batch processing pipelines.
// AI Prompts: Generate C# code using Aspose.Cells to apply a password, choose AES‑256 encryption, and save the workbook to a UNC path with error handling and local fallback. | Explain how Aspose.Cells embeds encryption settings in an XLSX file and why they survive network transfers. | Outline steps to verify UNC path accessibility before saving an encrypted workbook with Aspose.Cells. | Show how to configure custom encryption provider and key size in Aspose.Cells for maximum security.

using System;
using System.IO;
using Aspose.Cells;

// C# example that creates a Workbook, writes data, sets a password, applies AES‑128 encryption via SetEncryptionOptions, checks the UNC network folder, falls back to a local directory if needed, and saves the encrypted XLSX file so the protection remains intact.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add sample data to the first worksheet
            wb.Worksheets[0].Cells["A1"].PutValue("Sensitive data");

            // Set the password that encrypts the workbook
            wb.Settings.Password = "MySecretPassword";

            // Optional: specify stronger encryption options (AES 128‑bit)
            wb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // UNC path to the network share where the file will be saved
            string networkPath = @"\\ServerName\SharedFolder\EncryptedWorkbook.xlsx";

            // Determine if the target directory exists; if not, fall back to a local path
            string targetPath = networkPath;
            string targetDir = Path.GetDirectoryName(networkPath);

            if (string.IsNullOrEmpty(targetDir) || !Directory.Exists(targetDir))
            {
                // Fallback to a local folder (current directory)
                targetPath = Path.Combine(Directory.GetCurrentDirectory(), "EncryptedWorkbook.xlsx");
                Console.WriteLine($"Network path not available. Saving to local path: {targetPath}");
            }

            // Save the workbook; the encryption settings are preserved in the saved file
            wb.Save(targetPath, SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
