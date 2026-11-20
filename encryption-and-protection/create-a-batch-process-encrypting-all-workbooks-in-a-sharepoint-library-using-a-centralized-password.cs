// Title: Batch encrypt all Excel workbooks in a SharePoint library using Aspose.Cells for .NET with a single central password
// Common Searches: C# Aspose.Cells code to encrypt every .xlsx and .xls file in a SharePoint folder with one password | How to apply password protection to multiple Excel workbooks stored in SharePoint using Aspose.Cells | Recursive workbook encryption in .NET for SharePoint document library with Aspose.Cells | Batch protect Excel files in a local folder representing a SharePoint library using Aspose.Cells SaveOptions | Encrypt all Excel workbooks in bulk with a common password using Aspose.Cells for .NET
// Tags: Aspose.Cells batch workbook encryption | central password Excel protection C# | recursive SharePoint Excel file processing Aspose.Cells | OoxmlSaveOptions password encryption .xlsx | XlsSaveOptions workbook password protection | overwrite original workbook with encrypted version Aspose.Cells

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The utility scans a specified folder (mirroring a SharePoint library), loads each .xlsx or .xls workbook with Aspose.Cells, sets a shared password via Workbook.Settings.Password, selects the appropriate SaveOptions (OoxmlSaveOptions for .xlsx, XlsSaveOptions for .xls), saves to a temporary file, and overwrites the original file, thereby batch‑encrypting all workbooks with a single central password.
class SharePointWorkbookEncryptor
{
    // Central password to encrypt all workbooks
    private const string EncryptionPassword = "YourCentralPassword";

    static void Main()
    {
        try
        {
            // Local folder containing workbooks to encrypt
            const string folderPath = @"C:\InputWorkbooks";

            // Verify folder exists to avoid FileNotFoundException
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Get all Excel files recursively
            var excelFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));

            foreach (var filePath in excelFiles)
            {
                try
                {
                    // Ensure the file still exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found (skipped): {filePath}");
                        continue;
                    }

                    // Load workbook from file
                    var workbook = new Workbook(filePath);

                    // Set workbook password for encryption
                    workbook.Settings.Password = EncryptionPassword;

                    // Choose appropriate save options based on file extension
                    SaveOptions saveOptions;
                    if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        // OOXML (XLSX) format options
                        saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
                    }
                    else // .xls
                    {
                        // Legacy Excel format options
                        saveOptions = new XlsSaveOptions();
                    }

                    // Save encrypted workbook to a temporary file
                    string tempFile = Path.GetTempFileName();
                    workbook.Save(tempFile, saveOptions);

                    // Overwrite original file with encrypted version
                    File.Copy(tempFile, filePath, true);
                    File.Delete(tempFile);

                    Console.WriteLine($"Encrypted: {filePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Encryption of all workbooks completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
