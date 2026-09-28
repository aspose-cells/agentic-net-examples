// Title: Encrypt multiple ODS spreadsheets with unique SHA‑256‑derived passwords using Aspose.Cells for .NET
// AI Prompts: Create a C# console application that scans a folder for .ods files, generates a 16‑character password from each file name using SHA‑256, applies the password as the workbook opening password via Aspose.Cells, and saves the protected files to a separate output directory. | Extend the batch ODS encryption tool to write a CSV log containing the source file name, the derived password, and the output path while preserving the encrypted workbook creation.
// Common Searches: how to batch encrypt ODS files with a different password for each file in C# using Aspose.Cells | C# generate 16 character password from filename SHA256 for spreadsheet protection | set opening password on OpenDocument spreadsheet with Aspose.Cells .NET | automate encryption of multiple ODS workbooks and save to another folder in .NET
// Tags: batch ODS encryption Aspose.Cells | SHA256 filename password generation C# | set workbook opening password OpenDocument | automated spreadsheet protection .NET | process multiple ODS files Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The sample scans a directory for .ods files, derives a 16‑character password from each file name using SHA‑256, loads each workbook with Aspose.Cells, assigns the password as the opening password, and saves the encrypted workbook to a target folder, enabling per‑file protection in bulk.
class OdsBatchEncryptor
{
    // Compute a password from the file name using SHA256 and take the first 16 characters of the hex string.
    private static string DerivePassword(string fileName)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(fileName));
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hashBytes)
                sb.Append(b.ToString("x2"));
            // Use first 16 characters (8 bytes) as the password.
            return sb.ToString().Substring(0, 16);
        }
    }

    static void Main(string[] args)
    {
        // Input folder containing ODS files.
        string inputFolder = @"C:\InputOds";
        // Output folder where encrypted ODS files will be saved.
        string outputFolder = @"C:\EncryptedOds";

        // Ensure output directory exists.
        Directory.CreateDirectory(outputFolder);

        try
        {
            // Process each ODS file in the input folder.
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.ods"))
            {
                try
                {
                    // Derive a unique password from the file name (without extension).
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                    string password = DerivePassword(fileNameWithoutExt);

                    // Verify the source file exists before loading.
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}");
                        continue;
                    }

                    // Load the workbook.
                    Workbook workbook = new Workbook(filePath);

                    // Apply password protection (opening password) to the workbook.
                    workbook.Settings.Password = password;

                    // Set up ODS save options (no password property needed).
                    OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.Ods);

                    // Determine the output file path.
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the encrypted workbook.
                    workbook.Save(outputPath, saveOptions);

                    Console.WriteLine($"Encrypted '{Path.GetFileName(filePath)}' with password '{password}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch encryption completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
