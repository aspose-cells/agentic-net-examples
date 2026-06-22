using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sensitive data");

                // Set the password required to open the workbook
                workbook.Settings.Password = "StrongPassword123";

                // Specify stronger encryption (AES 128‑bit) for Excel 2007+ files
                workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                // UNC path where the file will be saved
                string networkPath = @"\\MyServer\SharedFolder\EncryptedWorkbook.xlsx";

                // Ensure the target directory exists; if not, fallback to a local folder
                string targetDirectory = Path.GetDirectoryName(networkPath);
                if (string.IsNullOrEmpty(targetDirectory) || !Directory.Exists(targetDirectory))
                {
                    // Fallback to the current directory
                    networkPath = Path.Combine(Directory.GetCurrentDirectory(), "EncryptedWorkbook.xlsx");
                    targetDirectory = Directory.GetCurrentDirectory();
                }

                // Save the workbook with the password and encryption settings
                workbook.Save(networkPath, SaveFormat.Xlsx);

                // Verify that the file is encrypted by loading it with the password
                if (File.Exists(networkPath))
                {
                    LoadOptions loadOptions = new LoadOptions { Password = "StrongPassword123" };
                    Workbook loadedWorkbook = new Workbook(networkPath, loadOptions);
                    Console.WriteLine("Workbook loaded successfully. Cell A1 value: " +
                                      loadedWorkbook.Worksheets[0].Cells["A1"].Value);
                }
                else
                {
                    Console.WriteLine($"File not found after save operation: {networkPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}