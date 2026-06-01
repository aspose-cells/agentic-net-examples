using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookDecryptor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect the first argument to be the CSV file path.
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

            // Read each line of the CSV (format: workbookPath,password)
            foreach (var line in File.ReadLines(csvPath))
            {
                // Skip empty lines
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Split by comma – assume no commas inside fields
                var parts = line.Split(new[] { ',' }, 2);
                if (parts.Length != 2)
                {
                    Console.WriteLine($"Invalid line (expected two columns): {line}");
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
                    // Load the workbook with the supplied password
                    LoadOptions loadOptions = new LoadOptions();
                    loadOptions.Password = password;

                    Workbook wb = new Workbook(workbookPath, loadOptions);

                    // Remove encryption by clearing the password property
                    wb.Settings.Password = null;

                    // Determine output path – same folder with "_decrypted" suffix
                    string directory = Path.GetDirectoryName(workbookPath);
                    string filenameWithoutExt = Path.GetFileNameWithoutExtension(workbookPath);
                    string extension = Path.GetExtension(workbookPath);
                    string outputPath = Path.Combine(directory, $"{filenameWithoutExt}_decrypted{extension}");

                    // Save the unprotected workbook
                    wb.Save(outputPath);

                    Console.WriteLine($"Decrypted workbook saved to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process '{workbookPath}': {ex.Message}");
                }
            }
        }
    }
}