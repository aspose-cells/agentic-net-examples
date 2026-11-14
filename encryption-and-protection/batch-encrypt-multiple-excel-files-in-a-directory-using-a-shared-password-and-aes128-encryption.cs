// Title: Encrypt all .xlsx workbooks in a folder with a shared password using Aspose.Cells for .NET (AES‑128)
// Common Searches: C# code to apply the same opening password to every .xlsx file in a directory with Aspose.Cells | How to batch encrypt Excel files using Aspose.Cells and AES-128 in .NET | Aspose.Cells encrypt multiple workbooks programmatically with a common password | Loop through a folder and set Workbook.Settings.Password for each Excel file in C# | Encrypt a batch of Excel spreadsheets with a shared password using Aspose.Cells .NET API
// Tags: mass Excel workbook encryption Aspose.Cells | shared password AES-128 for .xlsx files | Workbook.Settings.Password C# | process Excel files in a directory .NET | programmatic Excel file protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example scans a specified folder for .xlsx files, loads each workbook with Aspose.Cells, assigns a common opening password via the Workbook.Settings.Password property (which applies AES‑128 encryption), saves the workbook back to its original location, and logs progress and any errors to the console.
class BatchEncryptExcel
{
    static void Main()
    {
        // Directory containing the Excel files to encrypt
        string folderPath = @"C:\ExcelFiles";

        // Verify that the directory exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Shared password for all files
        string password = "SharedPassword123";

        // Get all .xlsx files in the directory
        string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string filePath in excelFiles)
        {
            try
            {
                // Ensure the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(filePath);

                // Set opening password (encrypts the file)
                workbook.Settings.Password = password;

                // Save the workbook (overwrites the original file)
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Encryption completed for {0} files.", excelFiles.Length);
    }
}
