// Title: Encrypt multiple Excel workbooks in parallel with Aspose.Cells using SHA‑256 derived passwords (C#)
// Common Searches: C# encrypt a folder of .xlsx files concurrently with Aspose.Cells | generate a unique password per workbook using SHA256 in .NET | parallel processing of Excel workbook encryption using Aspose.Cells | how to apply password protection to each Excel file in a multithreaded C# application | use Parallel.ForEach to encrypt multiple workbooks with different passwords
// Tags: Aspose.Cells concurrent workbook encryption | SHA256 derived password for Excel protection | C# multithreaded XLSX encryption | unique per‑file password Aspose.Cells | parallel file processing with Parallel.ForEach .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample enumerates all .xlsx files in a source directory, creates a distinct 16‑character password for each workbook by hashing a seed (file name plus timestamp) with SHA‑256, loads the workbook via Aspose.Cells, assigns the password through wb.Settings.Password, and saves the encrypted copy to a target folder. Encryption of each workbook runs concurrently using Parallel.ForEach, providing fast, thread‑safe processing of multiple Excel files.
class WorkbookEncryptor
{
    // Derive a password from a given seed using SHA256 and return a 16‑character string
    private static string DerivePassword(string seed)
    {
        using (SHA256 sha = SHA256.Create())
        {
            byte[] hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(seed));
            // Convert to Base64 and take first 16 characters for a manageable password length
            return Convert.ToBase64String(hashBytes).Substring(0, 16);
        }
    }

    // Encrypt a single workbook and save it to the destination folder
    private static void EncryptWorkbook(string sourcePath, string destFolder)
    {
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

        // Load the workbook (Aspose.Cells)
        Workbook wb = new Workbook(sourcePath);

        // Derive a unique password for this workbook (e.g., based on file name and current ticks)
        string passwordSeed = Path.GetFileNameWithoutExtension(sourcePath) + DateTime.UtcNow.Ticks;
        string password = DerivePassword(passwordSeed);

        // Apply password protection (Aspose.Cells encryption)
        wb.Settings.Password = password;

        // Prepare destination path
        string destPath = Path.Combine(destFolder, Path.GetFileName(sourcePath));

        // Save the encrypted workbook
        wb.Save(destPath, SaveFormat.Xlsx);
    }

    public static void Main()
    {
        // Folder containing source workbooks
        string sourceFolder = @"C:\Workbooks\Source";

        // Folder where encrypted workbooks will be saved
        string encryptedFolder = @"C:\Workbooks\Encrypted";

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder does not exist: {sourceFolder}");
            return;
        }

        // Ensure destination folder exists
        Directory.CreateDirectory(encryptedFolder);

        // Gather all Excel files to process
        List<string> workbookFiles = new List<string>(Directory.GetFiles(sourceFolder, "*.xlsx"));

        if (workbookFiles.Count == 0)
        {
            Console.WriteLine("No Excel files found to encrypt.");
            return;
        }

        // Process files in parallel, each on its own thread
        Parallel.ForEach(workbookFiles, sourcePath =>
        {
            try
            {
                EncryptWorkbook(sourcePath, encryptedFolder);
                Console.WriteLine($"Encrypted: {Path.GetFileName(sourcePath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to encrypt {Path.GetFileName(sourcePath)}: {ex.Message}");
            }
        });

        Console.WriteLine("All workbooks processed.");
    }
}
