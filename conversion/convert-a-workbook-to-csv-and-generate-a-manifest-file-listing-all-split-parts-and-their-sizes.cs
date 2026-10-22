// Title: Convert each worksheet of an Excel workbook to separate CSV files and create a manifest with file names and sizes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# split workbook into individual CSV files per worksheet | how to generate a manifest file with CSV part sizes using Aspose.Cells .NET | save each Excel sheet as CSV and record file size in a text file with Aspose.Cells | C# code to export worksheets to CSV and create a size manifest using Aspose.Cells | Aspose.Cells example for converting workbook to multiple CSV files and writing a manifest
// Tags: worksheet to CSV conversion Aspose.Cells | split workbook into CSV parts .NET | generate CSV manifest file Aspose.Cells | record CSV file size C# | per-sheet CSV export Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an XLSX workbook, iterates through its worksheets, creates a temporary workbook containing only the current sheet, saves each sheet as an individual CSV file in a designated folder, captures each CSV file's byte size, and writes a manifest.txt listing the CSV filenames with their corresponding sizes.
class WorkbookToCsvSplitter
{
    static void Main(string[] args)
    {
        // Input workbook path
        string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at path '{inputPath}'.");
            return;
        }

        // Output directory for CSV parts
        string outputDir = "CsvParts";
        Directory.CreateDirectory(outputDir);

        // List to hold manifest entries
        List<string> manifestLines = new List<string>();

        try
        {
            // Load the source workbook
            Workbook sourceWb = new Workbook(inputPath);

            // Iterate through each worksheet and save it as an individual CSV file
            foreach (Worksheet sheet in sourceWb.Worksheets)
            {
                // Create a new workbook that will contain only the current sheet
                Workbook partWb = new Workbook();

                // Remove the default empty sheet created by the constructor
                partWb.Worksheets.Clear();

                // Add a new (empty) worksheet to the part workbook and obtain its reference
                int newSheetIndex = partWb.Worksheets.Add();
                Worksheet newSheet = partWb.Worksheets[newSheetIndex];

                // Copy the current sheet into the new worksheet
                sheet.Copy(newSheet);

                // Build the CSV file name (e.g., Sheet1.csv)
                string csvFileName = $"{sheet.Name}.csv";
                string csvFilePath = Path.Combine(outputDir, csvFileName);

                // Save the part workbook as CSV
                partWb.Save(csvFilePath, SaveFormat.Csv);

                // Get the file size in bytes
                long fileSize = new FileInfo(csvFilePath).Length;

                // Add entry to manifest (format: FileName,SizeInBytes)
                manifestLines.Add($"{csvFileName},{fileSize}");
            }

            // Write the manifest file
            string manifestPath = Path.Combine(outputDir, "manifest.txt");
            File.WriteAllLines(manifestPath, manifestLines);

            Console.WriteLine("Workbook split into CSV parts and manifest generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
