// Title: C# – Split an Excel workbook into per‑worksheet CSV files and generate a size manifest using Aspose.Cells
// Description: Loads an Excel workbook, iterates through each worksheet, saves the active sheet as a CSV file with a sanitized name, records the file size, and writes a manifest.txt that lists every CSV file and its byte size in a chosen output folder. Fully compatible with .NET.
// Keywords: Aspose.Cells CSV split C# | export Excel worksheets to CSV .NET | generate CSV manifest file | worksheet to CSV conversion Aspose | file size listing for split CSV | sanitize file names Excel to CSV | batch CSV export Aspose.Cells
// Common Searches: How to split an Excel file into separate CSVs with Aspose.Cells C# | Create a manifest of CSV files and sizes after splitting workbook | Save each worksheet as CSV using Aspose.Cells .NET | C# code to export Excel sheets to individual CSV files | Aspose.Cells generate CSV size report
// Developer Intent: The developer needs to break a multi‑sheet Excel workbook into individual CSV files—one per worksheet—and produce a manifest that records each CSV file name and its byte size.
// Use Cases: Automate data‑pipeline preparation by exporting each sheet to CSV while tracking file sizes for downstream validation. | Provide per‑sheet CSV extracts for reporting systems that require separate files and an audit log. | Support large multi‑sheet workbooks in batch jobs, creating a manifest for storage‑quota monitoring.
// AI Prompts: Write C# code with Aspose.Cells that splits a workbook into CSV files and creates a manifest containing file names and sizes, handling invalid filename characters. | Show how to add worksheet index and a timestamp to each CSV file name and manifest entry. | Suggest robust error handling and logging for the CSV split and manifest generation process.

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCsvSplit
{
    // Loads an Excel workbook, iterates through each worksheet, saves the active sheet as a CSV file with a sanitized name, records the file size, and writes a manifest.txt that lists every CSV file and its byte size in a chosen output folder. Fully compatible with .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Directory where split CSV files and the manifest will be stored
            string outputDir = "SplitCsvOutput";
            Directory.CreateDirectory(outputDir);

            // Load the workbook using the constructor that accepts a file path (load rule)
            Workbook workbook = new Workbook(sourcePath);

            // List to hold manifest entries
            List<string> manifestLines = new List<string>();

            // Iterate through each worksheet, save it as an individual CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Activate the current worksheet so that Save with CSV format writes only this sheet
                workbook.Worksheets.ActiveSheetIndex = i;

                // Build CSV file name based on worksheet name
                string sheetName = workbook.Worksheets[i].Name;
                // Replace any invalid filename characters
                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    sheetName = sheetName.Replace(c, '_');
                }
                string csvFileName = $"{sheetName}.csv";
                string csvPath = Path.Combine(outputDir, csvFileName);

                // Save the active sheet as CSV (save rule)
                workbook.Save(csvPath, SaveFormat.Csv);

                // Get file size in bytes
                long fileSize = new FileInfo(csvPath).Length;

                // Add entry to manifest
                manifestLines.Add($"{csvFileName}, {fileSize} bytes");
            }

            // Write the manifest file
            string manifestPath = Path.Combine(outputDir, "manifest.txt");
            File.WriteAllLines(manifestPath, manifestLines);

            Console.WriteLine($"Workbook split into {workbook.Worksheets.Count} CSV files.");
            Console.WriteLine($"Manifest generated at: {manifestPath}");
        }
    }
}
