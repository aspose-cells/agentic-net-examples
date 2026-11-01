// Title: Encrypt an Excel workbook with a password and save it to a UNC network share using Aspose.Cells for .NET, with local fallback
// Common Searches: Aspose.Cells C# set password on workbook and save to UNC path | Save password‑protected Excel file to network share with Aspose.Cells | Check network share availability before saving encrypted workbook in .NET | Use temporary local folder when UNC share is not reachable with Aspose.Cells | Ensure workbook encryption is retained when saving as XLSX using Aspose.Cells
// Tags: Aspose.Cells workbook password encryption | save encrypted Excel to network share | C# verify UNC directory existence | temporary folder fallback Aspose.Cells | preserve XLSX encryption on save

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionExampleApp
{
    // The example creates a new Workbook, adds data, assigns a password via workbook.Settings.Password, verifies whether a specified UNC network share exists, and if not redirects the save to a temporary local folder. It then saves the workbook as XLSX, preserving the encryption, and reports success or errors.
    class WorkbookEncryptionExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // Set the encryption password for the workbook
                workbook.Settings.Password = "MySecurePassword123";

                // Define the network share path where the encrypted workbook will be saved
                string networkPath = @"\\ServerName\ShareFolder\EncryptedWorkbook.xlsx";

                // Determine if the network directory exists; if not, fall back to a local path
                string targetPath = networkPath;
                string networkDir = Path.GetDirectoryName(networkPath);
                if (string.IsNullOrEmpty(networkDir) || !Directory.Exists(networkDir))
                {
                    string tempDir = Path.GetTempPath();
                    targetPath = Path.Combine(tempDir, "EncryptedWorkbook.xlsx");
                    Console.WriteLine($"Network path not found. Saving to local path: {targetPath}");
                }

                // Save the workbook preserving the encryption
                workbook.Save(targetPath, SaveFormat.Xlsx);

                Console.WriteLine("Workbook saved with encryption successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
