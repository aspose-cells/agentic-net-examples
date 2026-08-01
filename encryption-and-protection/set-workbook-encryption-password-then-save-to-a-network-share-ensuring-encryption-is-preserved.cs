// Title: Encrypt Excel Workbook with Password and Save to UNC Network Share using Aspose.Cells for .NET
// Description: Creates a new Workbook, adds data, applies a password (which encrypts the file), optionally sets strong AES encryption, validates a UNC network share path, falls back to a local folder if needed, and saves the encrypted .xlsx file.
// Keywords: Aspose.Cells password protection | C# encrypt Excel workbook | save encrypted Excel to UNC path | network share Excel file Aspose | strong encryption Aspose.Cells | fallback to local folder C# | Excel file security .NET
// Common Searches: Aspose.Cells set workbook password C# | save encrypted Excel to network share UNC | C# Excel encryption with fallback directory | how to apply strong AES encryption using Aspose.Cells | check network share existence before saving Excel
// Developer Intent: Secure an Excel workbook with a password (and optional strong encryption) and reliably store it on a network share, using a fallback to a local directory when the share is unavailable.
// Use Cases: Protect confidential financial reports before publishing them to a shared drive. | Automate generation of encrypted audit logs saved directly to a central server. | Ensure reliable saving of encrypted spreadsheets in environments with intermittent network share access.
// AI Prompts: Write C# code with Aspose.Cells that encrypts a workbook using a 256‑bit AES password and saves it to a UNC path, including error handling for inaccessible shares. | Provide an example that checks a network share's availability, applies password protection, and falls back to a local "Output" folder when saving an Excel file with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a new Workbook, adds data, applies a password (which encrypts the file), optionally sets strong AES encryption, validates a UNC network share path, falls back to a local folder if needed, and saves the encrypted .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sensitive information");

            // Set a password – this encrypts the workbook so a password is required to open it
            workbook.Settings.Password = "StrongPassword123";

            // (Optional) Specify stronger encryption options (relevant for older Excel formats)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Define the UNC path of the network share where the file will be saved
            string networkSharePath = @"\\MyServer\SharedFolder\EncryptedWorkbook.xlsx";

            // Verify that the target directory exists; if not, fall back to a local folder
            string targetDirectory = Path.GetDirectoryName(networkSharePath);
            if (string.IsNullOrEmpty(targetDirectory) || !Directory.Exists(targetDirectory))
            {
                // Use a local "Output" folder relative to the current directory
                targetDirectory = Path.Combine(Environment.CurrentDirectory, "Output");
                Directory.CreateDirectory(targetDirectory);
                networkSharePath = Path.Combine(targetDirectory, "EncryptedWorkbook.xlsx");
            }

            // Save the encrypted workbook to the determined location
            workbook.Save(networkSharePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {networkSharePath}");
        }
        catch (Exception ex)
        {
            // Log or display the error details
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
