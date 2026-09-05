// Title: Batch encrypt Excel workbooks stored in a SharePoint library with a single password using Aspose.Cells for .NET
// AI Prompts: Write C# code that enumerates all .xls, .xlsx, and .xlsm files in a SharePoint document library and encrypts each workbook with a common password using Aspose.Cells. | Add logging to the sample so that each workbook's encryption result (file path, success/failure, error details) is written to a CSV report. | Enhance the utility to detect read‑only files, skip them, and continue processing the remaining workbooks without terminating the batch. | Create a PowerShell script that invokes the compiled C# tool, passing the SharePoint site URL, library name, and the centralized password as arguments.
// Common Searches: how to apply the same password to multiple Excel files in a SharePoint folder using Aspose.Cells C# | C# program to encrypt all .xlsx files in a SharePoint document library in bulk | batch protect Excel workbooks stored on SharePoint with Aspose.Cells API | automate workbook password protection for files in SharePoint using .NET | encrypt Excel workbooks recursively from a network share with Aspose.Cells
// Tags: Aspose.Cells bulk workbook password protection | C# encrypt Excel files in SharePoint library | centralized password for multiple workbooks | batch processing Excel files with Aspose.Cells | save encrypted workbook as Xlsx using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionDemo
{
    // The example walks through a given folder (including subfolders), identifies Excel files (.xls, .xlsx, .xlsm), loads each workbook with Aspose.Cells, assigns a predefined password via workbook.Settings.Password, and saves the file back in Xlsx format, overwriting the original while handling per‑file errors. The logic can be adapted to a SharePoint library by replacing the local directory enumeration with SharePoint file retrieval.
    class Program
    {
        // Centralized password for all workbooks
        private const string WorkbookPassword = "YourSecurePassword";

        // Folder containing Excel files to encrypt
        private const string InputFolderPath = @"C:\ExcelFiles";

        static void Main()
        {
            try
            {
                if (!Directory.Exists(InputFolderPath))
                {
                    Console.WriteLine($"Input folder does not exist: {InputFolderPath}");
                    return;
                }

                // Process all Excel files in the folder and subfolders
                var excelFiles = Directory.GetFiles(InputFolderPath, "*.*", SearchOption.AllDirectories);
                foreach (var filePath in excelFiles)
                {
                    if (IsExcelFile(filePath))
                    {
                        EncryptWorkbook(filePath);
                    }
                }

                Console.WriteLine("Encryption process completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Determines whether the file has an Excel extension
        private static bool IsExcelFile(string filePath)
        {
            var ext = Path.GetExtension(filePath);
            return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                   ext.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
                   ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
        }

        // Encrypts a single workbook with the predefined password
        private static void EncryptWorkbook(string filePath)
        {
            try
            {
                Console.WriteLine($"Encrypting: {filePath}");

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load workbook using Aspose.Cells
                var workbook = new Workbook(filePath);

                // Set the password for the workbook
                workbook.Settings.Password = WorkbookPassword;

                // Save the encrypted workbook (overwrite original)
                workbook.Save(filePath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to encrypt '{filePath}': {ex.Message}");
            }
        }
    }
}
