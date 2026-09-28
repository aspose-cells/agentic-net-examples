// Title: Measure and log encryption time for each Excel workbook in a folder using Aspose.Cells for .NET
// AI Prompts: Create a C# console app that scans a directory for .xlsx files, encrypts each workbook with a given password using Aspose.Cells, saves the encrypted files to an output folder, and prints the filename with elapsed milliseconds to the console. | Update the script to write the filename, encryption duration, and any error messages to a CSV log file instead of the console, creating the log file if it does not exist. | Extend the program to handle .xls, .xlsx, and .xlsm files, record the encryption time for each format, and display a summary of total files processed and average encryption time.
// Common Searches: C# how to batch encrypt Excel files with Aspose.Cells and capture processing time per file | Aspose.Cells encrypt multiple workbooks and log duration in a .NET console application | measure performance of password protection for Excel workbooks using Aspose.Cells C# | log encryption time for each .xlsx file while saving with Aspose.Cells | automate Excel file password protection and timing with Aspose.Cells for .NET
// Tags: Aspose.Cells workbook password protection C# | encrypt multiple Excel workbooks .NET | measure workbook encryption latency | record encryption duration per file C# | save protected .xlsx using Aspose.Cells

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example iterates through all .xlsx files in a specified input folder, loads each workbook with Aspose.Cells, applies a password to encrypt it, saves the encrypted copy to an output folder, and writes the elapsed milliseconds for each file to the console.
class Program
{
    static void Main(string[] args)
    {
        // Folder containing the original Excel files
        string inputFolder = @"C:\Data\Input";
        // Folder where encrypted files will be saved
        string outputFolder = @"C:\Data\Output";
        // Password to encrypt each workbook
        string password = "Secret123";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            string fileName = Path.GetFileName(filePath);
            string outputPath = Path.Combine(outputFolder, fileName);

            try
            {
                // Verify the file exists before loading
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Start timing the encryption operation
                Stopwatch sw = Stopwatch.StartNew();

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Apply password protection (encryption)
                workbook.Settings.Password = password;

                // Save the encrypted workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                // Stop timing
                sw.Stop();

                // Log the time taken for this file
                Console.WriteLine($"{fileName}: {sw.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }
    }
}
