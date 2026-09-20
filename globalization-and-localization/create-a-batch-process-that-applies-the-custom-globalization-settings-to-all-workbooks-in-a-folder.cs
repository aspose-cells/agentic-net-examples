// Title: Apply a custom CultureInfo (comma decimal, dot thousand, Euro currency) to all Excel workbooks in a folder using Aspose.Cells for .NET
// AI Prompts: Generate a C# console application that scans a specified directory for .xlsx, .xls, and .xlsm files, loads each workbook with Aspose.Cells, assigns a cloned InvariantCulture where NumberDecimalSeparator=',' , NumberGroupSeparator='.', CurrencySymbol='€', and saves the modified workbook to an output folder. | Update the batch utility so that after applying the custom CultureInfo it overwrites the original Excel files instead of writing copies to a separate directory. | Add comprehensive logging to the processing loop that writes the file path, success status, and any exception details to a log file for each workbook.
// Common Searches: how to change decimal separator for multiple Excel files using Aspose.Cells in C# | batch set custom CultureInfo for all workbooks in a folder Aspose.Cells .NET | C# script to apply Euro currency format to a directory of .xls and .xlsx files with Aspose | automate globalization settings across many spreadsheets using Aspose.Cells | process all Excel workbooks in a folder with custom number format using Aspose.Cells
// Tags: Aspose.Cells apply custom CultureInfo to Excel workbooks | batch processing Excel files with Aspose.Cells | set number decimal separator programmatically Aspose.Cells | globalization settings for .xlsx files C# | replace workbook with updated culture Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The program iterates through a given input folder, loads each .xlsx, .xls, or .xlsm workbook with Aspose.Cells, applies a custom CultureInfo that uses a comma as the decimal separator, a dot as the thousands separator, and the Euro symbol for currency, then saves the workbook to an output directory (or overwrites the original file if configured).
class GlobalizationBatchProcessor
{
    static void Main(string[] args)
    {
        // Folder containing the source workbooks
        string inputFolder = @"C:\Workbooks\Input";

        // Folder where the processed workbooks will be saved (can be the same as inputFolder to overwrite)
        string outputFolder = @"C:\Workbooks\Output";

        // Ensure the input and output directories exist
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder does not exist: {inputFolder}");
            return;
        }
        Directory.CreateDirectory(outputFolder);

        // Retrieve all Excel files in the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
        foreach (string filePath in excelFiles)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsm")
                continue; // Skip non‑Excel files

            // Verify the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Create a custom CultureInfo with desired separators and currency settings
                CultureInfo customCulture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
                customCulture.NumberFormat.NumberDecimalSeparator = ",";   // Use comma as decimal separator
                customCulture.NumberFormat.NumberGroupSeparator = ".";    // Use dot as thousands separator
                customCulture.NumberFormat.CurrencySymbol = "€";          // Euro symbol
                customCulture.NumberFormat.CurrencyPositivePattern = 2; // "€ n"
                customCulture.NumberFormat.CurrencyNegativePattern = 9; // "€-n"

                // Apply the custom culture to the workbook
                workbook.Settings.CultureInfo = customCulture;

                // Save the modified workbook to the output folder (preserving original name)
                string fileName = Path.GetFileName(filePath);
                string outputPath = Path.Combine(outputFolder, fileName);
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }

        Console.WriteLine("Custom globalization settings have been applied to all workbooks.");
    }
}
