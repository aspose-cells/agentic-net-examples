// Title: C# – Batch encrypt Excel files in a folder with a shared password using AES‑128 (Aspose.Cells)
// Description: A C# console example that scans a directory, loads each .xls/.xlsx/.xlsm/.xlsb workbook with Aspose.Cells, sets a common password, applies AES‑128 encryption via StrongCryptographicProvider, and saves the protected files to a target folder while preserving original names and formats.
// Keywords: Aspose.Cells | C# batch encrypt Excel | AES-128 Excel encryption | shared password Excel | StrongCryptographicProvider | encrypt multiple workbooks | Excel file protection .NET | programmatic Excel security | folder encryption Aspose | Excel encryption example
// Common Searches: batch encrypt Excel files C# | apply same password to multiple Excel workbooks Aspose.Cells | AES-128 encryption for Excel using .NET | how to protect all Excel files in a folder programmatically | set StrongCryptographicProvider encryption for Excel files | encrypt xls and xlsx files with Aspose.Cells
// Developer Intent: Programmatically protect every Excel workbook in a directory with a single password using AES‑128 encryption.
// Use Cases: Secure a batch of financial reports before archiving them on a shared drive. | Apply company‑wide password protection to generated spreadsheets before distributing them to external partners. | Automate encryption of exported data files in a nightly build to satisfy compliance and data‑privacy policies.
// AI Prompts: Generate C# code that uses Aspose.Cells to encrypt all .xlsx files in a folder with a shared password and AES‑256 encryption. | Show how to modify the sample to skip already encrypted workbooks, log success or failure for each file, and write the log to a CSV file. | Create a PowerShell script that runs the compiled batch‑encryption executable, passes source and destination paths as parameters, and captures the console output.

using System;
using System.IO;
using Aspose.Cells;

// A C# console example that scans a directory, loads each .xls/.xlsx/.xlsm/.xlsb workbook with Aspose.Cells, sets a common password, applies AES‑128 encryption via StrongCryptographicProvider, and saves the protected files to a target folder while preserving original names and formats.
class BatchEncryptExcel
{
    static void Main()
    {
        // Directory containing the original Excel files
        string sourceDirectory = @"C:\InputExcelFiles";

        // Directory where encrypted files will be saved
        string destinationDirectory = @"C:\EncryptedExcelFiles";

        // Shared password for all files
        const string sharedPassword = "MySharedPassword123";

        // Ensure the destination directory exists
        Directory.CreateDirectory(destinationDirectory);

        // Process each Excel file in the source directory
        foreach (string filePath in Directory.GetFiles(sourceDirectory, "*.*", SearchOption.TopDirectoryOnly))
        {
            // Consider only Excel formats (xls, xlsx, xlsm, xlsb, etc.)
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm" && extension != ".xlsb")
                continue;

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Set the password that protects the workbook
            workbook.Settings.Password = sharedPassword;

            // Apply AES‑128 encryption (StrongCryptographicProvider with 128‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

            // Build the output file path (same name, same extension)
            string destPath = Path.Combine(destinationDirectory, Path.GetFileName(filePath));

            // Save the encrypted workbook (extension determines format)
            workbook.Save(destPath);
        }

        Console.WriteLine("Batch encryption completed.");
    }
}
