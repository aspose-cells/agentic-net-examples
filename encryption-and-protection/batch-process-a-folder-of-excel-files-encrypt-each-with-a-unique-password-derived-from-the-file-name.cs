// Title: Encrypt a batch of Excel .xlsx files with passwords derived from their filenames using Aspose.Cells for .NET
// AI Prompts: Create a C# console application that scans a folder, loads each .xlsx workbook with Aspose.Cells, assigns the file name (without extension) as the opening password, and saves the workbook back to the same file. | Enhance the program to generate a SHA‑256 hash of each file name and use that hash string as the workbook password for stronger security. | Adapt the solution to also handle .xls and .xlsm files, preserving their original formats while applying the same filename‑based password logic.
// Common Searches: Aspose.Cells C# encrypt multiple Excel files in a folder using filename as password | How to set opening password for each .xlsx file in a directory with Aspose.Cells | Batch protect Excel workbooks with unique passwords derived from file names in .NET | C# loop through folder and apply password protection to Excel files using Aspose.Cells API
// Tags: batch workbook encryption Aspose.Cells | set opening password programmatically C# | derive password from filename Aspose.Cells | iterate over .xlsx files directory .NET | apply password protection to Excel files

using System;
using System.IO;
using Aspose.Cells;

// The program iterates through a specified directory, loads every .xlsx workbook with Aspose.Cells, sets the workbook's opening password to the file name (without extension), and saves the file, effectively encrypting each Excel file with a unique password based on its name.
class Program
{
    static void Main(string[] args)
    {
        // Path to the folder containing Excel files
        string folderPath = @"C:\Path\To\Folder";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Process all .xlsx files in the folder (add other extensions if needed)
        string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");

        foreach (string filePath in excelFiles)
        {
            try
            {
                // Ensure the file still exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Derive a simple password from the file name (without extension)
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string password = fileName; // Replace with a hash for stronger security if needed

                // Load the workbook from the file
                Workbook workbook = new Workbook(filePath);

                // Set the password that will be required to open the workbook
                workbook.Settings.Password = password;

                // Save the workbook back to the same file (overwrites the original)
                workbook.Save(filePath, SaveFormat.Xlsx);
                Console.WriteLine($"Processed: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
