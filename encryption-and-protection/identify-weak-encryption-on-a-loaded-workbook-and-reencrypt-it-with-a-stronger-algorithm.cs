// Title: C# – Detect Weak Excel Encryption and Upgrade to AES‑256 with Aspose.Cells
// Description: Loads an Excel workbook that may be protected with a weak password, checks Workbook.Settings.IsEncrypted, assigns a stronger password, applies AES‑256 encryption via SetEncryptionOptions, and saves the file. Includes fallback loading without a password and demonstrates how to re‑encrypt a workbook in one step.
// Keywords: Aspose.Cells | C# | .NET | Excel encryption | AES-256 | re‑encrypt workbook | weak password | Workbook.Settings.IsEncrypted | SetEncryptionOptions | upgrade Excel protection | password‑protected Excel file
// Common Searches: How to change Excel file encryption to AES‑256 using Aspose.Cells | Detect if an .xlsx is encrypted and re‑save with a stronger password in C# | Upgrade weak Excel workbook protection with Aspose.Cells .NET | Set strong encryption for a password‑protected workbook Aspose.Cells
// Developer Intent: Upgrade an existing Excel file from weak or unknown encryption to strong AES‑256 protection using Aspose.Cells for .NET.
// Use Cases: A security audit tool that scans Excel files, identifies weak encryption, and re‑encrypts them with a new strong password. | Automated batch process that opens password‑protected workbooks, applies AES‑256 encryption, and saves the updated files. | Legacy application migration where old Excel files encrypted with outdated algorithms need to meet modern compliance standards.
// AI Prompts: Generate C# code with Aspose.Cells to open a password‑protected .xlsx, verify Workbook.Settings.IsEncrypted, set a new password, apply AES‑256 encryption via SetEncryptionOptions, and save the workbook. | Explain the relationship between Workbook.Settings.IsEncrypted, Workbook.Settings.Password, and Workbook.SetEncryptionOptions when strengthening Excel file protection. | Provide a step‑by‑step guide for a PowerShell script that uses Aspose.Cells to batch re‑encrypt all Excel files in a folder from weak encryption to AES‑256.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook that may be protected with a weak password, checks Workbook.Settings.IsEncrypted, assigns a stronger password, applies AES‑256 encryption via SetEncryptionOptions, and saves the file. Includes fallback loading without a password and demonstrates how to re‑encrypt a workbook in one step.
class ReEncryptWorkbook
{
    static void Main()
    {
        // Path to the workbook that may be weakly encrypted
        string inputPath = "weak_encrypted.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Password used to open the existing workbook (if it is protected)
        string existingPassword = "weakpwd";

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook with the provided password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = existingPassword
            };
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException ex)
        {
            // If loading fails due to an invalid password, try loading without a password
            Console.WriteLine($"Failed to open with password: {ex.Message}");
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"Unable to load workbook: {innerEx.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error while loading workbook: {ex.Message}");
            return;
        }

        // Determine whether the workbook is encrypted
        bool isEncrypted = workbook.Settings.IsEncrypted;

        // If it is encrypted (or even if not), apply stronger encryption
        if (isEncrypted)
        {
            // Define a new strong password
            string newPassword = "StrongPwd123!";

            // Apply the new password
            workbook.Settings.Password = newPassword;

            // Set strong encryption: AES with 256‑bit key
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);
        }

        // Save the workbook with the stronger encryption
        string outputPath = "strong_encrypted.xlsx";

        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
