// Title: Generate a CSV report of password‑protected Excel files in a directory using Aspose.Cells for .NET
// Common Searches: how to list encrypted Excel workbooks in a folder with Aspose.Cells C# | C# code to scan directories for password protected .xlsx files using Aspose.Cells | create CSV file showing which Excel files require a password in .NET | detect Excel workbook encryption without opening the file using Aspose.Cells | Aspose.Cells load options to check for password protection in bulk
// Tags: Aspose.Cells detect encrypted workbook | CSV password status report C# | directory scan for protected Excel files | LoadOptions password detection Aspose.Cells | CellsException handling for encrypted Excel

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace ExcelPasswordStatusReport
{
    // A C# utility that recursively scans a specified folder for Excel files, attempts to open each workbook with Aspose.Cells without a password, catches CellsException to determine if the file is password‑protected, and writes a UTF‑8 CSV report listing the full file path and a Yes/No flag indicating encryption status, with proper CSV escaping and error handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Directory to scan – change as needed or pass via args
                string targetDirectory = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

                if (!Directory.Exists(targetDirectory))
                {
                    Console.WriteLine($"Directory does not exist: {targetDirectory}");
                    return;
                }

                // Output report file
                string reportPath = Path.Combine(targetDirectory, "PasswordStatusReport.csv");

                // Prepare CSV header
                var sb = new StringBuilder();
                sb.AppendLine("FilePath,IsEncrypted");

                // Supported Excel extensions
                string[] extensions = new[] { ".xls", ".xlsx", ".xlsm", ".xlsb" };

                // Enumerate files recursively
                foreach (string file in Directory.EnumerateFiles(targetDirectory, "*.*", SearchOption.AllDirectories))
                {
                    if (Array.IndexOf(extensions, Path.GetExtension(file).ToLower()) < 0)
                        continue; // Skip non‑Excel files

                    if (!File.Exists(file))
                        continue; // Safety check

                    bool isEncrypted = false;

                    try
                    {
                        // Attempt to load without a password; will throw if protected
                        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
                        Workbook wb = new Workbook(file, loadOptions);
                        // Load succeeded – not encrypted
                        isEncrypted = false;
                    }
                    catch (CellsException ex)
                    {
                        // Check message for password requirement
                        if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            isEncrypted = true;
                        }
                        else
                        {
                            // Other CellsException – treat as not encrypted
                            isEncrypted = false;
                        }
                    }
                    catch (Exception)
                    {
                        // Unexpected exception – treat as not encrypted to keep processing
                        isEncrypted = false;
                    }

                    // Append result to CSV
                    sb.AppendLine($"{EscapeCsv(file)},{(isEncrypted ? "Yes" : "No")}");
                }

                // Write the report to disk
                try
                {
                    File.WriteAllText(reportPath, sb.ToString(), Encoding.UTF8);
                    Console.WriteLine($"Password status report generated at: {reportPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write report: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Helper to escape CSV fields that may contain commas or quotes
        private static string EscapeCsv(string field)
        {
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            {
                field = field.Replace("\"", "\"\"");
                return $"\"{field}\"";
            }
            return field;
        }
    }
}
