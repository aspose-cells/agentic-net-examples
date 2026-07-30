// Title: C# batch encrypt ODS files with filename‑derived passwords using Aspose.Cells
// Description: Scans a folder for *.ods workbooks, loads each with Aspose.Cells OdsLoadOptions, creates a unique 16‑character password from the SHA‑256 hash of the file name, applies it via workbook.Settings.Password, and saves the encrypted file with OdsSaveOptions to a target directory.
// Keywords: Aspose.Cells ODS encryption C# | batch ODS password protection | derive password from filename SHA256 | C# encrypt multiple ODS files | Aspose.Cells OdsLoadOptions example | workbook.Settings.Password | ODS file security .NET
// Common Searches: how to encrypt multiple ODS spreadsheets with Aspose.Cells | C# generate password from file name hash for ODS encryption | batch process ODS files and save encrypted versions | Aspose.Cells example encrypt ODS workbook | C# SHA256 password for ODS file protection
// Developer Intent: Encrypt every ODS workbook in a directory using a deterministic password derived from its file name.
// Use Cases: Automatically protect incoming ODS reports before archiving, using a reproducible password per file. | Secure a batch of ODS documents for compliance by applying per‑file encryption in a .NET service. | Integrate ODS encryption into a CI/CD pipeline to safeguard generated spreadsheets prior to deployment.
// AI Prompts: Write C# code that encrypts all .ods files in a folder with Aspose.Cells, using a 16‑character password derived from the SHA‑256 hash of each file name. | Show how to replace SHA‑256 with MD5 for password generation in the batch ODS encryption example. | Explain how to decrypt the ODS files produced by this program using Aspose.Cells and the filename‑based password.

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace BatchOdsEncryption
{
    // Scans a folder for *.ods workbooks, loads each with Aspose.Cells OdsLoadOptions, creates a unique 16‑character password from the SHA‑256 hash of the file name, applies it via workbook.Settings.Password, and saves the encrypted file with OdsSaveOptions to a target directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing source ODS files
            string inputFolder = @"C:\Ods\Input";
            // Folder where encrypted ODS files will be saved
            string outputFolder = @"C:\Ods\Encrypted";

            // Ensure the input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder '{inputFolder}' does not exist. Please create it and add .ods files.");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Process each .ods file in the input folder
            foreach (string inputFilePath in Directory.GetFiles(inputFolder, "*.ods"))
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(inputFilePath))
                    {
                        Console.WriteLine($"File not found: {inputFilePath}");
                        continue;
                    }

                    // Derive a unique password from the file name (using SHA256 hash)
                    string fileName = Path.GetFileNameWithoutExtension(inputFilePath);
                    string password = DerivePasswordFromFileName(fileName);

                    // Load the workbook with default ODS load options
                    OdsLoadOptions loadOptions = new OdsLoadOptions();
                    Workbook workbook = new Workbook(inputFilePath, loadOptions);

                    // Set the workbook password (encryption)
                    workbook.Settings.Password = password;

                    // Prepare ODS save options
                    OdsSaveOptions saveOptions = new OdsSaveOptions();

                    // Save the encrypted workbook to the output folder
                    string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(inputFilePath));
                    workbook.Save(outputFilePath, saveOptions);

                    Console.WriteLine($"Encrypted '{Path.GetFileName(inputFilePath)}' with password '{password}'. Saved to '{outputFilePath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{inputFilePath}': {ex.Message}");
                }
            }
        }

        // Generates a password string from the SHA256 hash of the file name.
        // Takes the first 16 characters of the hex representation for brevity.
        private static string DerivePasswordFromFileName(string fileName)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(fileName));
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                    sb.Append(b.ToString("x2"));
                // Use first 16 characters (8 bytes) as the password
                return sb.ToString().Substring(0, 16);
            }
        }
    }
}
