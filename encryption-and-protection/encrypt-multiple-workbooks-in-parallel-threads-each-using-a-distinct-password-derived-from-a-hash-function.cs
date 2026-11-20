// Title: Encrypt multiple Excel workbooks concurrently with Aspose.Cells for .NET using SHA‑256 derived passwords
// Common Searches: how to encrypt several .xlsx files simultaneously using Aspose.Cells in C# | generate unique workbook opening passwords from SHA256 hash with Aspose.Cells | parallel processing of Excel files for password protection using Parallel.ForEach and Aspose.Cells | set workbook.Settings.Password programmatically in Aspose.Cells .NET | save encrypted Excel workbook with OoxmlSaveOptions in Aspose.Cells
// Tags: parallel encryption of Excel workbooks Aspose.Cells | SHA256 derived opening password Aspose.Cells | Workbook.Settings.Password assignment C# | OoxmlSaveOptions encrypted .xlsx Aspose.Cells | Parallel.ForEach workbook protection .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads a list of .xlsx files, creates an output directory, and uses Parallel.ForEach to process each workbook on a separate thread. For each file a SHA‑256 hash of the file path is computed, the first 16 hex characters become the opening password (assigned to workbook.Settings.Password), and the workbook is saved as an encrypted .xlsx using OoxmlSaveOptions. Missing files and runtime errors are caught and logged.
class Program
{
    static void Main()
    {
        // List of source workbook file paths to encrypt
        var sourceFiles = new List<string>
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
        };

        // Ensure output directory exists
        string outputDir = "EncryptedWorkbooks";
        Directory.CreateDirectory(outputDir);

        // Encrypt each workbook in parallel
        Parallel.ForEach(sourceFiles, sourcePath =>
        {
            try
            {
                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(sourcePath);

                // Derive a unique password from a hash of the source path
                string password = DerivePassword(sourcePath);

                // Apply opening password protection
                workbook.Settings.Password = password;

                // Build the output file name
                string fileName = Path.GetFileNameWithoutExtension(sourcePath);
                string outputPath = Path.Combine(outputDir, $"{fileName}_encrypted.xlsx");

                // Save the encrypted workbook
                var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Encrypted workbook saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{sourcePath}': {ex.Message}");
            }
        });
    }

    // Generates a password by taking the first 16 hex characters of a SHA256 hash
    static string DerivePassword(string input)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(input));
            var sb = new StringBuilder();
            // Use first 8 bytes (16 hex chars) as password
            for (int i = 0; i < 8; i++)
            {
                sb.Append(hashBytes[i].ToString("x2"));
            }
            return sb.ToString();
        }
    }
}
