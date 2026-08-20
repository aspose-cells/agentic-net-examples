// Title: C# Bulk Decrypt Password‑Protected Excel Files from CSV Using Aspose.Cells
// Description: A console utility that reads a CSV file containing Excel workbook paths and their passwords, loads each protected workbook with Aspose.Cells LoadOptions, clears the password, and overwrites the original file while reporting success or errors.
// Keywords: Aspose.Cells | C# bulk Excel decryption | remove Excel password .NET | CSV driven workbook unlock | load protected workbook | overwrite Excel file | GitHub example | source code | API usage
// Common Searches: batch remove password from Excel using Aspose.Cells | C# program to decrypt multiple protected workbooks | read CSV of Excel passwords and unlock files | Aspose.Cells load workbook with password and save unprotected | automate Excel password removal .NET
// Developer Intent: Read a CSV of file paths and passwords, then programmatically decrypt each protected Excel workbook with Aspose.Cells.
// Use Cases: Automate nightly decryption of incoming password‑protected reports before data‑pipeline ingestion. | Validate a list of encrypted workbooks and produce unprotected copies for internal analysis. | Provide a bulk decryption tool for compliance teams to remove passwords from archived Excel files on shared drives.
// AI Prompts: Generate C# code that reads a CSV of Excel file paths and passwords, loads each workbook with Aspose.Cells, clears its password, and saves it back, handling missing files and malformed lines. | Suggest enhancements to log detailed errors and optionally save decrypted files to a separate output folder. | Explain how to preserve original file timestamps after decryption using Aspose.Cells. | Create a GitHub‑ready project structure for this bulk decryption tool, including README and CI configuration.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookDecryptionTool
{
    // A console utility that reads a CSV file containing Excel workbook paths and their passwords, loads each protected workbook with Aspose.Cells LoadOptions, clears the password, and overwrites the original file while reporting success or errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the CSV file containing "filePath,password" per line
            string csvPath = "workbooks.csv";

            // Ensure the CSV file exists
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
                return;
            }

            // Read all lines from the CSV
            string[] lines = File.ReadAllLines(csvPath);

            foreach (string line in lines)
            {
                // Skip empty lines
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Split line into file path and password (comma‑separated)
                string[] parts = line.Split(new[] { ',' }, 2);
                if (parts.Length != 2)
                {
                    Console.WriteLine($"Invalid line format (expected 'path,password'): {line}");
                    continue;
                }

                string filePath = parts[0].Trim();
                string password = parts[1].Trim();

                // Verify the workbook file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Workbook not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the protected workbook using LoadOptions.Password (rule)
                    LoadOptions loadOptions = new LoadOptions();
                    loadOptions.Password = password;
                    Workbook workbook = new Workbook(filePath, loadOptions); // create/load rule

                    // Remove the password protection (set to null)
                    workbook.Settings.Password = null; // workbooksettings.password rule

                    // Save the workbook, overwriting the original file (save rule)
                    workbook.Save(filePath);

                    Console.WriteLine($"Successfully decrypted: {filePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to decrypt {filePath}: {ex.Message}");
                }
            }
        }
    }
}
