// Title: Decrypt multiple password‑protected Excel workbooks in bulk using a CSV of file paths and passwords with Aspose.Cells for .NET
// Common Searches: C# program to read a CSV of Excel file paths and passwords and decrypt each workbook with Aspose.Cells | how to remove password protection from many .xlsx files using Aspose.Cells LoadOptions | batch process encrypted Excel files in .NET by supplying passwords from a text file | automate decryption of password‑protected workbooks with an Aspose.Cells console application | load encrypted XLSX with password from CSV and save unprotected using Aspose.Cells
// Tags: Aspose.Cells bulk workbook decryption | load encrypted XLSX with password C# | CSV driven Excel password removal | programmatic workbook protection clearing | Aspose.Cells LoadOptions password parameter

using System;
using System.IO;
using Aspose.Cells;

// A console utility reads a CSV where each line contains an Excel file path and its password, loads each encrypted workbook using Aspose.Cells LoadOptions, clears the workbook's password, and saves the file back unencrypted, handling missing files and runtime errors.
class WorkbookDecryptor
{
    static void Main(string[] args)
    {
        // Expect the first argument to be the path of the CSV file containing "filePath,password" per line.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: WorkbookDecryptor <csvFilePath>");
            return;
        }

        string csvPath = args[0];
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"CSV file not found: {csvPath}");
            return;
        }

        // Read all non‑empty lines from the CSV.
        var lines = File.ReadAllLines(csvPath);
        foreach (var rawLine in lines)
        {
            if (string.IsNullOrWhiteSpace(rawLine))
                continue;

            // Simple CSV split on first comma (file path may contain commas if quoted – not handled here).
            var parts = rawLine.Split(new[] { ',' }, 2);
            if (parts.Length != 2)
            {
                Console.WriteLine($"Invalid line (expected two columns): {rawLine}");
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
                // Load the encrypted workbook using the supplied password.
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                var workbook = new Workbook(workbookPath, loadOptions);

                // Remove the password protection.
                workbook.Settings.Password = string.Empty;

                // Save the workbook back, overwriting the original file (now unencrypted).
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
