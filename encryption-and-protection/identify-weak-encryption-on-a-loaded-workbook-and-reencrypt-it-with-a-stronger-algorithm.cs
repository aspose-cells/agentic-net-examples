// Title: Re‑encrypt an Excel workbook to AES‑256 with Aspose.Cells for .NET (C#)
// Description: Loads a password‑protected workbook, checks if it is encrypted, upgrades the protection to AES‑256 using Aspose.Cells, retains the original password, and saves the file to a new location. Demonstrates how to replace weak encryption with a strong algorithm in C#.
// Keywords: Aspose.Cells re‑encrypt workbook | AES 256 Excel encryption C# | upgrade weak Excel encryption | SetEncryptionOptions Aspose.Cells | load encrypted .xlsx password | strong cryptographic provider | Excel file security .NET
// Common Searches: How to change Excel file encryption to AES‑256 using Aspose.Cells | Detect and upgrade weak encryption in a .xlsx with C# | Re‑encrypt password‑protected workbook Aspose.Cells .NET | Increase Excel encryption strength programmatically | Replace 128‑bit encryption with 256‑bit in Aspose.Cells
// Developer Intent: Replace a workbook’s weak encryption with AES‑256 while keeping the same password.
// Use Cases: Modernize legacy spreadsheets that were encrypted with 128‑bit keys before distribution. | Validate incoming Excel files in an automated pipeline and enforce AES‑256 protection. | Automatically re‑encrypt user‑uploaded spreadsheets to satisfy GDPR, HIPAA, or other compliance standards.
// AI Prompts: Generate C# code using Aspose.Cells to open an encrypted .xlsx, verify its encryption status, and save it with AES‑256 preserving the original password. | Provide a step‑by‑step tutorial for detecting weak encryption in an Excel workbook and upgrading it to a strong provider with Aspose.Cells for .NET. | Explain how to handle a workbook that is not encrypted before applying new encryption options in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads a password‑protected workbook, checks if it is encrypted, upgrades the protection to AES‑256 using Aspose.Cells, retains the original password, and saves the file to a new location. Demonstrates how to replace weak encryption with a strong algorithm in C#.
class ReEncryptWorkbook
{
    static void Main()
    {
        // Paths for the source (weakly encrypted) and destination (strongly encrypted) files
        string sourcePath = "weak_encrypted.xlsx";
        string destinationPath = "strong_encrypted.xlsx";

        // Password that protects the source workbook
        string password = "oldPassword";

        try
        {
            // Ensure the source file exists; if not, create a simple workbook and apply encryption
            if (!File.Exists(sourcePath))
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample data");

                // Apply encryption (using strong provider with 128‑bit key as a placeholder)
                tempWb.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
                tempWb.Settings.Password = password;
                tempWb.Save(sourcePath);
            }

            // Load the workbook using the existing password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Verify that the workbook is encrypted before re‑encrypting
            if (workbook.Settings.IsEncrypted)
            {
                // Apply a stronger encryption algorithm (AES 256‑bit)
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
                // Preserve the original password for opening the file later
                workbook.Settings.Password = password;
            }

            // Save the workbook with the new, stronger encryption
            workbook.Save(destinationPath);
            Console.WriteLine("Workbook re‑encrypted successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
