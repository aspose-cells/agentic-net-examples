using System;
using System.IO;
using System.Security;
using Aspose.Cells;

namespace SharePointWorkbookEncryption
{
    class Program
    {
        // Central password used to encrypt every workbook
        private const string CentralPassword = "YourCentralPassword";

        static void Main(string[] args)
        {
            try
            {
                // Path to the folder containing Excel files to encrypt
                string folderPath = args.Length > 0 ? args[0] : @"C:\ExcelFiles";

                if (!Directory.Exists(folderPath))
                {
                    Console.WriteLine($"Folder not found: {folderPath}");
                    return;
                }

                ProcessFolder(folderPath, CentralPassword);
                Console.WriteLine("Processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <summary>
        /// Encrypts all Excel workbooks in the specified folder using Aspose.Cells.
        /// </summary>
        private static void ProcessFolder(string folderPath, string password)
        {
            foreach (string filePath in Directory.GetFiles(folderPath))
            {
                string fileName = Path.GetFileName(filePath);

                // Process only Excel workbook files
                if (!IsExcelFile(fileName))
                    continue;

                try
                {
                    // Load workbook from file
                    Workbook workbook = new Workbook(filePath);

                    // Apply password protection (encryption)
                    workbook.Settings.Password = password;

                    // Determine the appropriate SaveFormat based on original extension
                    SaveFormat format = GetSaveFormat(fileName);

                    // Save the encrypted workbook back to the same file (overwrite)
                    workbook.Save(filePath, format);

                    Console.WriteLine($"Encrypted: {fileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process {fileName}: {ex.Message}");
                }
            }
        }

        // Helper to check if a file is an Excel workbook
        private static bool IsExcelFile(string fileName)
        {
            string ext = Path.GetExtension(fileName).ToLowerInvariant();
            return ext == ".xlsx" || ext == ".xls" || ext == ".xlsm" || ext == ".xlsb" || ext == ".ods";
        }

        // Helper to map file extension to Aspose.Cells SaveFormat
        private static SaveFormat GetSaveFormat(string fileName)
        {
            string ext = Path.GetExtension(fileName).ToLowerInvariant();
            return ext switch
            {
                ".xlsx" => SaveFormat.Xlsx,
                ".xls" => SaveFormat.Excel97To2003,
                ".xlsm" => SaveFormat.Xlsm,
                ".xlsb" => SaveFormat.Xlsb,
                ".ods" => SaveFormat.ODS,
                _ => SaveFormat.Xlsx,
            };
        }
    }
}