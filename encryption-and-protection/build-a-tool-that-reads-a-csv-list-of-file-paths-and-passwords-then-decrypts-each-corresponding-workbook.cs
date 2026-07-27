using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPasswordDecryptor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the CSV file containing "filePath,password" per line
            string csvPath = "files_and_passwords.csv";

            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
                return;
            }

            // Read all non‑empty lines from the CSV
            string[] lines = File.ReadAllLines(csvPath);
            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue; // skip empty lines

                // Expecting two columns separated by a comma
                string[] parts = line.Split(new[] { ',' }, 2);
                if (parts.Length != 2)
                {
                    Console.WriteLine($"Invalid line format (expected 'path,password'): {line}");
                    continue;
                }

                string filePath = parts[0].Trim();
                string password = parts[1].Trim();

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Workbook not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook with the provided password
                    LoadOptions loadOptions = new LoadOptions();
                    loadOptions.Password = password; // rule: LoadOptions.Password

                    Workbook wb = new Workbook(filePath, loadOptions); // rule: Workbook(string, LoadOptions)

                    // Remove the password protection
                    wb.Settings.Password = null; // rule: WorkbookSettings.Password

                    // Save the workbook back to the same file (overwrites the encrypted version)
                    wb.Save(filePath); // rule: Workbook.Save(string)

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