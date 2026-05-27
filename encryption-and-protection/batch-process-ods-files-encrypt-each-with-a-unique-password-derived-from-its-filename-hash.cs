using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Ods;

class BatchEncryptOds
{
    // Derive a password from the file name using SHA256 and return a hex string
    private static string GetPasswordFromFileName(string filePath)
    {
        string fileName = Path.GetFileNameWithoutExtension(filePath);
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(fileName));
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hashBytes)
                sb.Append(b.ToString("x2"));
            return sb.ToString(); // 64‑character hex password
        }
    }

    static void Main()
    {
        // Folder containing ODS files to encrypt
        string inputFolder = @"C:\OdsFiles";

        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Get all .ods files in the folder (non‑recursive)
        string[] odsFiles = Directory.GetFiles(inputFolder, "*.ods", SearchOption.TopDirectoryOnly);

        foreach (string odsPath in odsFiles)
        {
            try
            {
                if (!File.Exists(odsPath))
                {
                    Console.WriteLine($"File not found, skipping: {odsPath}");
                    continue;
                }

                // Generate a unique password based on the file name
                string password = GetPasswordFromFileName(odsPath);

                // Load the ODS workbook
                Workbook workbook = new Workbook(odsPath);

                // Set the workbook encryption password
                workbook.Settings.Password = password;

                // Create ODS save options
                OdsSaveOptions saveOptions = new OdsSaveOptions
                {
                    GeneratorType = OdsGeneratorType.LibreOffice
                };

                // Save the workbook back, overwriting the original file
                workbook.Save(odsPath, saveOptions);

                Console.WriteLine($"Encrypted '{Path.GetFileName(odsPath)}' with password: {password}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{odsPath}': {ex.Message}");
            }
        }
    }
}