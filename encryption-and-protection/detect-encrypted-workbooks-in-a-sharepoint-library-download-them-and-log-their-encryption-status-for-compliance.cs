// Title: Detect password‑protected Excel workbooks in a SharePoint library and create a compliance log with Aspose.Cells for .NET
// Common Searches: C# code to scan a SharePoint folder for password‑protected Excel files using Aspose.Cells | How to log encryption status of Excel workbooks from a SharePoint document library in .NET | Identify encrypted .xlsx files in SharePoint with Aspose.Cells LoadOptions | Generate compliance report of encrypted Excel files stored on SharePoint using C# | Detect password‑protected workbooks in a SharePoint library without opening them
// Tags: Aspose.Cells load workbook without password | scan SharePoint folder for encrypted Excel workbooks | generate encryption compliance log .NET | handle CellsException for password‑protected workbooks | detect password‑protected .xlsx files using C#

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a specified directory that represents a SharePoint library, iterates through .xls, .xlsx, and .xlsm files, and attempts to open each workbook with Aspose.Cells without supplying a password. A CellsException indicates the workbook is password‑protected (encrypted). The program records the file name, timestamp, and encryption status (Encrypted or Not Encrypted) to a plain‑text compliance log.
class SharePointEncryptedWorkbookDetector
{
    static void Main()
    {
        // Folder containing Excel files to scan (replace with actual path)
        string folderPath = "InputFiles";

        // Path for the compliance log
        string logFilePath = "EncryptionLog.txt";

        // Verify the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Open the log file (overwrite any existing content)
        using (StreamWriter logWriter = new StreamWriter(logFilePath, false))
        {
            // Retrieve all Excel files in the folder and subfolders
            string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);
            foreach (string filePath in files)
            {
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                    continue; // Skip non‑Excel files

                string fileName = Path.GetFileName(filePath);

                // Ensure the file exists before processing
                if (!File.Exists(filePath))
                {
                    logWriter.WriteLine($"{DateTime.Now}: {fileName} - File not found.");
                    continue;
                }

                bool isEncrypted = false;

                try
                {
                    // Attempt to load the workbook without a password
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        LoadOptions loadOptions = new LoadOptions();
                        Workbook workbook = new Workbook(fs, loadOptions);
                        // If loaded successfully, the file is not password‑protected
                        isEncrypted = false;
                    }
                }
                catch (CellsException)
                {
                    // Any CellsException during load is treated as an encrypted (password‑protected) file
                    isEncrypted = true;
                }
                catch (Exception ex)
                {
                    // Log unexpected errors and continue with next file
                    logWriter.WriteLine($"{DateTime.Now}: {fileName} - Unexpected error: {ex.Message}");
                    continue;
                }

                string status = isEncrypted ? "Encrypted" : "Not Encrypted";
                logWriter.WriteLine($"{DateTime.Now}: {fileName} - {status}");
            }
        }

        Console.WriteLine($"Scanning completed. Log saved to {Path.GetFullPath(logFilePath)}");
    }
}
