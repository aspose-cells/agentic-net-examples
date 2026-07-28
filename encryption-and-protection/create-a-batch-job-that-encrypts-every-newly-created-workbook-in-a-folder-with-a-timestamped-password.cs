// Title: C# batch job to encrypt every Excel workbook in a folder with a timestamp‑generated password using Aspose.Cells
// Description: A C# console utility that scans a directory, loads each .xlsx, .xls, .xlsm, .xlsb, or .ods file with Aspose.Cells, creates a password from the current date‑time (yyyyMMddHHmmss), applies it via Workbook.Settings.Password, and saves the file, providing automatic, time‑based protection for newly created workbooks.
// Keywords: Aspose.Cells | C# encrypt Excel | batch workbook encryption | timestamp password | protect Excel files programmatically | folder scan encryption | Excel file security .NET | Workbook.Settings.Password | automated Excel protection | Aspose.Cells encryption example
// Common Searches: aspocells encrypt multiple workbooks c# | batch encrypt excel files with timestamp password | c# program to protect all Excel files in a folder | how to set password for Excel files using Aspose.Cells | automate Excel file encryption .NET
// Developer Intent: Automatically apply a unique, time‑based password to each Excel workbook placed in a specified folder.
// Use Cases: Secure daily generated reports before archiving by assigning a distinct timestamp password. | Integrate workbook protection into a CI/CD pipeline to satisfy compliance and data‑loss‑prevention policies. | Batch‑process a collection of spreadsheets for audit trails, ensuring each file has its own creation‑time password. | Provide on‑premises users with a simple script to lock all exported Excel files without manual intervention.
// AI Prompts: Generate C# code that watches a directory and encrypts any new Excel file with a password based on the current timestamp using Aspose.Cells. | Refactor the batch encryption program to log generated passwords to Azure Key Vault and skip files that are already password‑protected. | Create a PowerShell wrapper that invokes the C# utility and writes a CSV report of encrypted files and their timestamps. | Explain how to modify the example to use a custom password pattern (e.g., prefix + timestamp) while maintaining Aspose.Cells compatibility.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchEncryption
{
    // A C# console utility that scans a directory, loads each .xlsx, .xls, .xlsm, .xlsb, or .ods file with Aspose.Cells, creates a password from the current date‑time (yyyyMMddHHmmss), applies it via Workbook.Settings.Password, and saves the file, providing automatic, time‑based protection for newly created workbooks.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to encrypt
            string folderPath = @"C:\Workbooks";

            // Verify the folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Supported Excel file extensions
            string[] extensions = new[] { ".xlsx", ".xls", ".xlsm", ".xlsb", ".ods" };

            // Process each workbook file in the folder
            foreach (string filePath in Directory.GetFiles(folderPath))
            {
                // Skip files that are not Excel workbooks
                if (Array.IndexOf(extensions, Path.GetExtension(filePath).ToLower()) < 0)
                    continue;

                try
                {
                    // Load the existing workbook (load rule)
                    Workbook workbook = new Workbook(filePath);

                    // Generate a timestamped password (e.g., 20230727143055)
                    string timestampPassword = DateTime.Now.ToString("yyyyMMddHHmmss");

                    // Apply the password to the workbook (settings rule)
                    workbook.Settings.Password = timestampPassword;

                    // Save the workbook back to the same file (save rule)
                    workbook.Save(filePath);

                    Console.WriteLine($"Encrypted '{Path.GetFileName(filePath)}' with password: {timestampPassword}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to encrypt '{Path.GetFileName(filePath)}': {ex.Message}");
                }
            }
        }
    }
}
