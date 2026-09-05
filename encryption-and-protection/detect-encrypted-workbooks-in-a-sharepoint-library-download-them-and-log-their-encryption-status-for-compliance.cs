// Title: Use Aspose.Cells for .NET to detect password‑protected Excel workbooks in a local folder and log their encryption status
// AI Prompts: Write C# code that iterates through a directory, opens each .xls, .xlsx, and .xlsm file with Aspose.Cells, and determines whether the workbook is encrypted by catching CellsException. | Extend the program to export the file name together with a true/false encryption flag to a CSV file for compliance reporting. | Create a reusable method that receives a root folder path and returns a list of workbook paths that trigger a CellsException when opened without a password.
// Common Searches: how to programmatically identify password protected Excel files using Aspose.Cells in C# | C# scan folder for encrypted .xlsx files and generate compliance report | Aspose.Cells catch CellsException to check workbook encryption status | log encryption flag of multiple Excel workbooks with Aspose.Cells .NET | detect protected Excel workbooks in SharePoint downloaded folder using Aspose.Cells
// Tags: Aspose.Cells password protection check | C# enumerate Excel files for protection status | log Excel encryption flag to console | handle CellsException for protected workbooks | process .xls .xlsx .xlsm files with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace SharePointWorkbookEncryptionCheck
{
    // The sample enumerates all .xls, .xlsx, and .xlsm files in a specified folder, attempts to load each workbook with Aspose.Cells, treats a CellsException as an indication of password protection, and writes the file name with a boolean encryption flag to the console (or optionally to a CSV).
    class Program
    {
        static void Main(string[] args)
        {
            // Folder that contains Excel files to be checked.
            // Change this path to point to your actual directory.
            string inputFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "InputFiles");

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            string[] excelFiles;
            try
            {
                // Get all Excel files in the folder (including subfolders).
                excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.AllDirectories);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to enumerate files: {ex.Message}");
                return;
            }

            foreach (string filePath in excelFiles)
            {
                string extension = Path.GetExtension(filePath);
                if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                    !extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) &&
                    !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
                {
                    continue; // Skip non‑Excel files.
                }

                // Ensure the file exists before attempting to open it.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                bool isEncrypted = false;

                try
                {
                    // Attempt to load the workbook without a password.
                    // If the workbook is password‑protected, Aspose.Cells throws a CellsException.
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        Workbook wb = new Workbook(fs);
                        // Load succeeded – the workbook is not encrypted.
                        isEncrypted = false;
                    }
                }
                catch (CellsException)
                {
                    // Workbook is encrypted (password protected) or cannot be opened without a password.
                    isEncrypted = true;
                }
                catch (Exception ex)
                {
                    // Unexpected error – log and continue with the next file.
                    Console.WriteLine($"Error processing '{Path.GetFileName(filePath)}': {ex.Message}");
                    continue;
                }

                // Log the encryption status.
                Console.WriteLine($"{Path.GetFileName(filePath)}: Encrypted = {isEncrypted}");
            }
        }
    }
}
