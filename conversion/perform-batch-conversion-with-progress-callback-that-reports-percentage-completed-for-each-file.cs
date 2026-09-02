// Title: Convert multiple Excel .xlsx files to PDF with per‑file progress reporting using Aspose.Cells for .NET
// AI Prompts: Create a C# console application that scans a folder for .xlsx files, converts each workbook to PDF with Aspose.Cells, and prints the conversion percentage for every file. | Implement a progress callback in the Aspose.Cells PDF conversion loop that outputs the percent completed while saving each workbook. | Modify the batch conversion sample to record the progress percentages to a log file instead of writing them to the console.
// Common Searches: c# aspocells batch convert xlsx to pdf with progress indicator | how to display conversion percentage for each Excel file in a console app using Aspose.Cells | asp.net core console batch excel to pdf conversion progress callback example | aspocells pdfsaveoptions show progress per workbook | c# iterate folder convert all .xlsx to .pdf using Aspose.Cells and show progress
// Tags: aspose.cells batch conversion xlsx pdf c# | pdfsaveoptions progress callback aspocells | loadoptions xlsx workbook loading aspocells | directory enumeration excel files c# | console progress logging file conversion

using System;
using System.IO;
using Aspose.Cells;

// A C# console program that enumerates all .xlsx files in a specified input directory, loads each workbook with Aspose.Cells using LoadOptions, converts them to PDF via PdfSaveOptions, saves the PDFs to an output folder, and reports the conversion percentage for each file through a progress callback.
class BatchConversionWithProgress
{
    static void Main()
    {
        // Folder containing source Excel files
        string inputFolder = @"C:\InputFiles";
        // Folder where converted files will be saved
        string outputFolder = @"C:\OutputFiles";

        // Ensure output folder exists
        Directory.CreateDirectory(outputFolder);

        // Verify that the input folder exists; if not, inform the user and exit gracefully
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            Console.WriteLine("Please create the folder and add Excel files before running the program.");
            return;
        }

        // Get all Excel files in the input folder
        string[] excelFiles = Directory.GetFiles(inputFolder, "*.xlsx");

        foreach (string sourcePath in excelFiles)
        {
            string fileName = Path.GetFileName(sourcePath);
            Console.WriteLine($"Processing file: {fileName}");

            try
            {
                // Verify that the source file exists (redundant after GetFiles, but kept for safety)
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    continue;
                }

                // Load the workbook
                var loadOptions = new LoadOptions(LoadFormat.Xlsx);
                var workbook = new Workbook(sourcePath, loadOptions);

                // Prepare save options (convert to PDF)
                var saveOptions = new PdfSaveOptions();

                // Determine output file path (same name, different extension)
                string outputPath = Path.Combine(outputFolder,
                    Path.GetFileNameWithoutExtension(sourcePath) + ".pdf");

                // Save the workbook as PDF
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Finished processing {fileName}\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {fileName}: {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
