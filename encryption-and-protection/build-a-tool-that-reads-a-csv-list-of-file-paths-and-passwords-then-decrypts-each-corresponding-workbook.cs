// Title: Batch decrypt password‑protected Excel workbooks from a CSV list using Aspose.Cells for .NET (C#)
// Common Searches: how to batch decrypt encrypted Excel files with a CSV of passwords using Aspose.Cells in C# | C# code to read a CSV of workbook paths and passwords and remove Excel protection | automate removal of password protection from multiple Excel workbooks using Aspose.Cells LoadOptions | process a list of encrypted Excel files and save them unprotected with Aspose.Cells .NET | example of parsing CSV and decrypting Excel workbooks programmatically in C#
// Tags: mass decryption of Excel workbooks Aspose.Cells | CSV driven password removal for Excel files C# | Aspose.Cells LoadOptions password decryption | programmatic removal of Excel workbook protection .NET | process multiple encrypted workbooks using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// A C# console utility that accepts a CSV file containing Excel workbook paths and their passwords, loads each encrypted workbook with Aspose.Cells LoadOptions, clears the password, and saves the file back unprotected, handling missing files and errors while processing the list in bulk.
class Program
{
    static void Main(string[] args)
    {
        // Expect the CSV file path as the first argument
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: DecryptWorkbooks <list.csv>");
            return;
        }

        string csvPath = args[0];
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"CSV file not found: {csvPath}");
            return;
        }

        foreach (var line in File.ReadLines(csvPath))
        {
            // Skip empty lines
            if (string.IsNullOrWhiteSpace(line))
                continue;

            // Parse CSV line (path,password)
            var parts = ParseCsvLine(line);
            if (parts.Length < 2)
            {
                Console.WriteLine($"Invalid line (expected path,password): {line}");
                continue;
            }

            string workbookPath = parts[0];
            string password = parts[1];

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                continue;
            }

            try
            {
                // Load the encrypted workbook using the supplied password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                loadOptions.Password = password;
                Workbook wb = new Workbook(workbookPath, loadOptions);

                // Remove password protection
                wb.Settings.Password = string.Empty;

                // Save the workbook back to the same file (now unencrypted)
                wb.Save(workbookPath);
                Console.WriteLine($"Decrypted: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to decrypt {workbookPath}: {ex.Message}");
            }
        }
    }

    // Simple CSV parser that handles commas inside quoted fields
    static string[] ParseCsvLine(string line)
    {
        var fields = new System.Collections.Generic.List<string>();
        bool inQuotes = false;
        var current = new System.Text.StringBuilder();

        foreach (char c in line)
        {
            if (c == '\"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (c == ',' && !inQuotes)
            {
                fields.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(c);
            }
        }

        fields.Add(current.ToString());
        return fields.ToArray();
    }
}
