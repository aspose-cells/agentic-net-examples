// Title: Split Excel Workbook into CSV Files per Sheet and Create a Manifest with Sizes (Aspose.Cells C#)
// Description: Loads an Excel workbook with Aspose.Cells, saves each worksheet as an individual CSV using TxtSaveOptions (ExportAllSheets = false), records the file name and byte size, and writes a manifest.txt that lists all parts and their sizes.
// Keywords: Aspose.Cells CSV export | C# split workbook into CSV | generate manifest file .NET | TxtSaveOptions per sheet | Excel to CSV per worksheet | file size listing | Aspose.Cells example
// Common Searches: how to export each Excel sheet to separate CSV using Aspose.Cells | create manifest of CSV files with sizes in C# | Aspose.Cells save active worksheet as CSV | split workbook into multiple CSV files .NET | list CSV file sizes after export
// Developer Intent: Export every worksheet of an Excel workbook as a separate CSV file and produce a manifest that records each CSV's name and size.
// Use Cases: Automated data pipelines that require per‑sheet CSV outputs with a verification manifest. | Regulatory reporting where each sheet is delivered as a CSV and the manifest provides quick size validation. | CI/CD processes that break large workbooks into manageable CSV chunks and confirm output sizes before packaging.
// AI Prompts: Generate C# code with Aspose.Cells to split a workbook into CSV files per worksheet and create a manifest containing file names and byte sizes. | Show how to configure TxtSaveOptions to export only the active sheet to CSV and capture the resulting file size. | Explain how to extend the manifest format to include timestamps and MD5 hashes for each CSV part.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvSplit
{
    // Loads an Excel workbook with Aspose.Cells, saves each worksheet as an individual CSV using TxtSaveOptions (ExportAllSheets = false), records the file name and byte size, and writes a manifest.txt that lists all parts and their sizes.
    class Program
    {
        static void Main()
        {
            // Input Excel file path
            string inputPath = "input.xlsx";

            // Directory to store split CSV parts and manifest
            string outputDir = "output_parts";
            Directory.CreateDirectory(outputDir);

            // Load the workbook (using the constructor rule)
            Workbook workbook = new Workbook(inputPath);

            // List to hold manifest entries
            List<string> manifestLines = new List<string>();

            // Iterate through each worksheet and export it as an individual CSV file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                // Activate the current worksheet
                workbook.Worksheets.ActiveSheetIndex = i;

                // Prepare CSV save options: export only the active sheet
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
                csvOptions.ExportAllSheets = false; // export only the active sheet

                // Define part file name
                string partFileName = Path.Combine(outputDir, $"Part_{i + 1}.csv");

                // Save the active sheet as CSV (using Save(string, SaveOptions) rule)
                workbook.Save(partFileName, csvOptions);

                // Get file size in bytes
                long fileSize = new FileInfo(partFileName).Length;

                // Add entry to manifest
                manifestLines.Add($"{Path.GetFileName(partFileName)}\t{fileSize} bytes");
            }

            // Write manifest file
            string manifestPath = Path.Combine(outputDir, "manifest.txt");
            File.WriteAllLines(manifestPath, manifestLines);

            Console.WriteLine($"Workbook split into {workbook.Worksheets.Count} CSV parts.");
            Console.WriteLine($"Manifest generated at: {manifestPath}");
        }
    }
}
