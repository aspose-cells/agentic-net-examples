// Title: Batch encrypt ODS spreadsheets with filename‑derived passwords using Aspose.Cells for .NET
// Common Searches: C# code to protect ODS files with a password generated from the file name using Aspose.Cells | How to batch apply password protection to ODS workbooks in a folder with Aspose.Cells .NET | Generate SHA256 based password for each ODS file and encrypt it programmatically | Automate ODS workbook protection with unique passwords in a .NET console application | Aspose.Cells protect ODS spreadsheet when native encryption is unavailable
// Tags: batch ODS password protection Aspose.Cells | derive password from filename SHA256 C# | protect ODS workbook programmatically .NET | Aspose.Cells ODS save options encryption workaround | automated spreadsheet security .NET

using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Aspose.Cells;

// The example scans a directory for .ods files, creates a 16‑character password from each file name using SHA‑256, loads each workbook with Aspose.Cells, applies full protection with the derived password, and saves the workbook back in ODS format, overwriting the original files.
class OdsBatchEncryptor
{
    // Derive a password from the file name using SHA256 and take first 16 characters
    private static string DerivePassword(string fileName)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(fileName));
            StringBuilder sb = new StringBuilder();
            foreach (byte b in hashBytes)
                sb.Append(b.ToString("x2"));
            return sb.ToString().Substring(0, 16);
        }
    }

    static void Main(string[] args)
    {
        // Directory containing ODS files (change as needed)
        string sourceDirectory = @"C:\OdsFiles";

        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine($"Source directory does not exist: {sourceDirectory}");
            return;
        }

        // Get all .ods files in the directory
        string[] odsFiles = Directory.GetFiles(sourceDirectory, "*.ods", SearchOption.TopDirectoryOnly);

        foreach (string filePath in odsFiles)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Derive a unique password from the file name
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string password = DerivePassword(fileName);

                // Protect the workbook (since ODS encryption is not supported)
                workbook.Protect(ProtectionType.All, password);

                // Prepare save options for ODS format
                OdsSaveOptions saveOptions = new OdsSaveOptions();

                // Overwrite the original file with the protected version
                workbook.Save(filePath, saveOptions);

                Console.WriteLine($"Processed '{fileName}.ods' with password: {password}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch processing completed.");
    }
}
