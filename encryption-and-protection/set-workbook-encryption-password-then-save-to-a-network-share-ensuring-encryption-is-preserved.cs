// Title: Encrypt an Excel workbook with a password and save it to a UNC network share using Aspose.Cells for C#
// Common Searches: how to set a password on an Excel workbook with Aspose.Cells in C# | save a password-protected .xlsx file to a UNC path using Aspose.Cells | fallback to local temporary folder when network share is unavailable in Aspose.Cells C# | preserve workbook encryption while writing to a remote network location with Aspose.Cells | validate network directory before saving an Aspose.Cells workbook
// Tags: Aspose.Cells password protection workbook | save encrypted XLSX to UNC share C# | Workbook.Settings.Password Aspose.Cells | network path validation Aspose.Cells | temporary folder fallback Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Aspose.Cells Workbook, writes data to cell A1, applies a password via Workbook.Settings.Password, checks whether the target UNC network share exists, falls back to a local temporary file if needed, and saves the workbook as an encrypted .xlsx file using SaveFormat.Xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data
                workbook.Worksheets[0].Cells["A1"].PutValue("Hello, Aspose!");

                // Set password protection (required to open the file)
                workbook.Settings.Password = "StrongPassword123";

                // Define UNC network share path where the encrypted file will be saved
                string networkPath = @"\\FileServer\SharedFolder\EncryptedWorkbook.xlsx";

                // Determine a valid path; fallback to a local temp folder if the network share is unavailable
                string savePath = GetValidSavePath(networkPath);

                // Save the workbook; password protection is applied automatically
                workbook.Save(savePath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to: {savePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Returns a valid path; if the desired network directory does not exist, use a temporary local path
        private static string GetValidSavePath(string desiredPath)
        {
            try
            {
                string directory = Path.GetDirectoryName(desiredPath);
                if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                {
                    return desiredPath;
                }
            }
            catch
            {
                // Ignore any exception and fall back to temp path
            }

            // Fallback to a temporary file in the local temp folder
            return Path.Combine(Path.GetTempPath(), "EncryptedWorkbook.xlsx");
        }
    }
}
