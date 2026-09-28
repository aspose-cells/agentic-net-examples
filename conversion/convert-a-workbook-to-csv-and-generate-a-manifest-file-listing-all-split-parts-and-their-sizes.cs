// Title: Split an Excel workbook into individual CSV files per worksheet and generate a manifest.txt with each file’s byte size using Aspose.Cells for .NET
// AI Prompts: Create a C# console application that loads an .xlsx workbook with Aspose.Cells, iterates over all worksheets, saves each sheet as a separate CSV file, and writes a text file that records each CSV name and its size in bytes. | Write .NET code that uses Aspose.Cells to export every worksheet to CSV format and produces a summary file listing the generated CSV filenames alongside their byte counts.
// Common Searches: Aspose.Cells C# export each worksheet to its own CSV and capture file size | How to produce a size report for CSV files created from Excel sheets using Aspose.Cells | C# split Excel workbook into multiple CSV files and list their byte lengths | Generate a manifest of CSV parts from an Excel workbook with Aspose.Cells | Save worksheets as CSV and create a file size manifest in .NET
// Tags: Aspose.Cells worksheet to CSV conversion | CSV manifest generation with file sizes | SaveFormat.Csv per worksheet Aspose.Cells | Workbook split into multiple CSV files .NET | record CSV file byte length C#

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// The program loads 'input.xlsx', creates an 'output' folder, iterates through each worksheet, saves each as a CSV file named after the sheet, captures the file size in bytes, and writes a 'manifest.txt' that lists every CSV filename with its corresponding size.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Directory where CSV parts and the manifest will be saved
            string outputDir = "output";
            Directory.CreateDirectory(outputDir);

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Prepare a list to hold manifest entries
            var manifestLines = new List<string>();

            // Iterate through each worksheet and save it as an individual CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                try
                {
                    // Activate the current worksheet
                    workbook.Worksheets.ActiveSheetIndex = i;

                    // Build the CSV file name based on the worksheet name
                    string csvFileName = $"{workbook.Worksheets[i].Name}.csv";
                    string csvFilePath = Path.Combine(outputDir, csvFileName);

                    // Save the active worksheet as CSV (only the active sheet is exported)
                    workbook.Save(csvFilePath, SaveFormat.Csv);

                    // Get the size of the generated CSV file
                    long fileSize = new FileInfo(csvFilePath).Length;

                    // Add an entry to the manifest: file name and size in bytes
                    manifestLines.Add($"{csvFileName},{fileSize}");
                }
                catch (Exception sheetEx)
                {
                    Console.WriteLine($"Error processing worksheet '{workbook.Worksheets[i].Name}': {sheetEx.Message}");
                }
            }

            // Write the manifest file listing all CSV parts and their sizes
            string manifestPath = Path.Combine(outputDir, "manifest.txt");
            File.WriteAllLines(manifestPath, manifestLines);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
