// Title: Detect password‑protected Excel workbooks in a SharePoint document library and log their encryption status with Aspose.Cells for .NET
// Common Searches: C# how to identify password protected .xlsx files in a SharePoint library using Aspose.Cells | Aspose.Cells detect encrypted Excel workbook without providing a password in .NET | log encryption status of Excel files retrieved from a SharePoint folder | iterate through SharePoint document library Excel files and check for protection with Aspose.Cells | handle CellsException to determine if an Excel workbook is password protected in C#
// Tags: Aspose.Cells password protected workbook detection | C# enumerate SharePoint document library Excel files | log encrypted workbook status Aspose.Cells | CellsException handling for protected .xlsx | download Excel files from SharePoint using .NET | compliance check encrypted Excel workbooks

using System;
using System.IO;
using Aspose.Cells;

namespace SharePointWorkbookEncryptionCheck
{
    // The example scans a specified SharePoint folder (or local copy) for .xlsx files, attempts to open each workbook with Aspose.Cells without a password, catches a CellsException containing a password‑related message to flag the file as encrypted, and writes the filename together with its encryption status to the console for compliance reporting.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder path containing workbook files (replace with your actual path)
            string folderPath = @"C:\Workbooks";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Folder not found: {folderPath}");
                return;
            }

            // Iterate through all Excel files in the folder
            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
            {
                // Ensure the file exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                bool isEncrypted = false;

                try
                {
                    // Load workbook without providing a password.
                    // If the file is password‑protected, Aspose.Cells throws a CellsException.
                    using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        Workbook workbook = new Workbook(stream);
                    }
                }
                catch (CellsException ex)
                {
                    // Determine if the exception is due to password protection.
                    // Older Aspose.Cells versions may not expose ErrorCode, so inspect the message.
                    if (ex.Message != null &&
                        ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        isEncrypted = true;
                    }
                    else
                    {
                        Console.WriteLine($"CellsException processing {Path.GetFileName(filePath)}: {ex.Message}");
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    // Log other errors and continue processing
                    Console.WriteLine($"Error processing {Path.GetFileName(filePath)}: {ex.Message}");
                    continue;
                }

                // Log encryption status
                Console.WriteLine($"{Path.GetFileName(filePath)}: Encrypted = {isEncrypted}");
            }
        }
    }
}
