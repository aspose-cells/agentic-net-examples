// Title: Decrypt multiple password‑protected Excel workbooks from a CSV list using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that reads a CSV file of Excel file paths and passwords, opens each encrypted workbook with Aspose.Cells LoadOptions, clears the password, and saves the workbook back to disk. | Add robust error handling and create a log file that records the timestamp, file path, and success or failure of each decryption attempt. | Enhance the program to detect the workbook format (XLS, XLSX, XLSB) and set the appropriate LoadFormat when loading each encrypted file.
// Common Searches: how to batch remove passwords from Excel files using Aspose.Cells in C# | C# read CSV of file paths and passwords to decrypt encrypted workbooks | Aspose.Cells load encrypted XLSX with password from list and save unprotected | automate decryption of multiple protected Excel workbooks via command line | process CSV of Excel workbook locations and passwords with Aspose.Cells .NET
// Tags: batch decrypt Excel workbooks Aspose.Cells | load encrypted workbook with password C# | CSV-driven Excel file decryption .NET | remove password from XLSX using Aspose.Cells | automated Excel protection removal C#

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookDecryptor
{
    // A C# console tool reads a CSV where each line contains an Excel workbook path and its password, loads each encrypted workbook with Aspose.Cells LoadOptions, clears the password, and overwrites the original file with the decrypted version, while reporting success or failure for every entry.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect the first argument to be the path of the CSV file containing workbook paths and passwords.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the CSV file as the first argument.");
                return;
            }

            string csvPath = args[0];

            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"CSV file not found: {csvPath}");
                return;
            }

            // Read all lines from the CSV file.
            string[] lines = File.ReadAllLines(csvPath);

            foreach (string line in lines)
            {
                // Skip empty lines.
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Each line is expected to be: workbookPath,password
                string[] parts = line.Split(new[] { ',' }, 2);
                if (parts.Length != 2)
                {
                    Console.WriteLine($"Invalid line format (expected 'path,password'): {line}");
                    continue;
                }

                string workbookPath = parts[0].Trim();
                string password = parts[1].Trim();

                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook not found: {workbookPath}");
                    continue;
                }

                try
                {
                    // Load the encrypted workbook using the provided password.
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                    {
                        Password = password
                    };

                    Workbook workbook = new Workbook(workbookPath, loadOptions);

                    // Remove the password protection.
                    workbook.Settings.Password = null; // or string.Empty

                    // Overwrite the original file with the decrypted version.
                    workbook.Save(workbookPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Successfully decrypted: {workbookPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to decrypt {workbookPath}: {ex.Message}");
                }
            }
        }
    }
}
