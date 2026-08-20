// Title: Batch encrypt ODS files with filename‑derived SHA‑256 passwords using Aspose.Cells for .NET
// Description: A C# example that scans a folder for *.ods files, creates a SHA‑256 hash of each file name (without extension), converts the first 16 bytes to a hex string, sets this string as the workbook password, and saves the encrypted workbook with Aspose.Cells OdsSaveOptions to a target directory.
// Keywords: Aspose.Cells | C# ODS encryption | batch ODS password protection | filename hash password | SHA256 password generation | OdsSaveOptions | encrypt multiple ODS files | Aspose.Cells for .NET example | GitHub code sample | automated spreadsheet security
// Common Searches: encrypt all ODS files in a folder Aspose.Cells | C# generate password from file name for ODS encryption | batch protect ODS workbooks with Aspose.Cells | how to set workbook password programmatically .NET | example code for ODS encryption using Aspose
// Developer Intent: Automatically apply a unique, deterministic password to each ODS workbook in a directory and save the protected files.
// Use Cases: Secure confidential spreadsheets before archiving by using reproducible passwords derived from file names. | Automate protection of daily‑generated ODS reports, enabling later decryption with the same naming convention. | Integrate into CI/CD pipelines to encrypt ODS assets before distribution to partners or external systems.
// AI Prompts: Generate C# code that validates a user‑entered password against the filename‑derived password when opening an encrypted ODS file with Aspose.Cells. | Provide a method to decrypt an ODS workbook that was encrypted using the filename‑derived SHA‑256 password shown in the sample. | Suggest a secure logging strategy for the password‑derivation process that prevents exposing the hash or generated password.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Ods;

// A C# example that scans a folder for *.ods files, creates a SHA‑256 hash of each file name (without extension), converts the first 16 bytes to a hex string, sets this string as the workbook password, and saves the encrypted workbook with Aspose.Cells OdsSaveOptions to a target directory.
class BatchEncryptOds
{
    static void Main()
    {
        // Folder containing source ODS files
        string inputFolder = @"C:\Ods\Input";
        // Folder where encrypted ODS files will be saved
        string outputFolder = @"C:\Ods\Encrypted";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .ods file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.ods"))
        {
            try
            {
                // Derive a password from the file name (without extension) using SHA256
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string password = DerivePasswordFromName(fileName);

                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Set the workbook password
                workbook.Settings.Password = password;

                // Prepare ODS save options (optional: set generator type)
                OdsSaveOptions saveOptions = new OdsSaveOptions
                {
                    GeneratorType = OdsGeneratorType.LibreOffice
                };

                // Save the encrypted workbook to the output folder
                string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Encrypted '{Path.GetFileName(filePath)}' with password derived from its name.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }

    // Generates a password string from a given name using SHA256 and returns a hex representation
    private static string DerivePasswordFromName(string name)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(name));
            // Convert the first 16 bytes of the hash to a hex string for a reasonable password length
            StringBuilder sb = new StringBuilder(32);
            for (int i = 0; i < 16; i++)
            {
                sb.Append(hashBytes[i].ToString("x2"));
            }
            return sb.ToString();
        }
    }
}
