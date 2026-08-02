// Title: Batch encrypt Excel workbooks in a SharePoint library with a single password using Aspose.Cells for .NET
// Description: A C# console utility that scans a folder mapped to a SharePoint document library, loads each .xlsx file with Aspose.Cells, assigns a common password via Workbook.Settings.Password, and overwrites the file. Includes progress messages and robust error handling for large‑scale encryption.
// Keywords: Aspose.Cells | C# Excel encryption | SharePoint document library | batch workbook password | Workbook.Settings.Password | automated Excel protection | encrypt multiple .xlsx files | console app | file system traversal | compliance security
// Common Searches: encrypt all Excel files in SharePoint using Aspose.Cells | C# batch password protection for .xlsx | apply same password to multiple workbooks .NET | automate Excel encryption in SharePoint library | Aspose.Cells bulk workbook encryption example
// Developer Intent: Apply a uniform password to every Excel workbook stored in a SharePoint library.
// Use Cases: Protect confidential financial statements stored centrally before external sharing. | Run nightly job that secures newly uploaded Excel reports to meet regulatory standards. | Distribute corporate templates with a predefined password to enforce consistent access control.
// AI Prompts: Write a C# script that iterates through a SharePoint‑mapped directory and encrypts each .xlsx file with a given password using Aspose.Cells. | Show how to detect already‑protected workbooks and skip them during bulk encryption. | Add logging to the process that records successes, failures, and timestamps for audit purposes.

using System;
using System.IO;
using Aspose.Cells;

namespace SharePointWorkbookEncryption
{
    // A C# console utility that scans a folder mapped to a SharePoint document library, loads each .xlsx file with Aspose.Cells, assigns a common password via Workbook.Settings.Password, and overwrites the file. Includes progress messages and robust error handling for large‑scale encryption.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the workbooks to encrypt
            string folderPath = @"C:\Workbooks";               // change to your folder
            // Central password to apply to all workbooks
            string centralPassword = "YourCentralPassword";

            try
            {
                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                    return;
                }

                // Get all Excel files in the folder (including subfolders if needed)
                string[] files = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories);

                foreach (string filePath in files)
                {
                    try
                    {
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found, skipping: {filePath}");
                            continue;
                        }

                        // Load workbook from file
                        Workbook workbook = new Workbook(filePath);

                        // Apply password protection (encryption)
                        workbook.Settings.Password = centralPassword;

                        // Overwrite the original file with the encrypted version
                        workbook.Save(filePath, SaveFormat.Xlsx);

                        Console.WriteLine($"Encrypted: {filePath}");
                    }
                    catch (Exception exFile)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {exFile.Message}");
                    }
                }

                Console.WriteLine("Encryption of all workbooks completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
