// Title: Save a workbook as XLSX and verify that the generated package does not contain printer‑settings .bin files with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to load an existing workbook, save it as a new XLSX file, and iterate through the ZIP entries to detect any .bin files. | Create a .NET console application that reports whether printer‑settings binary files are present inside the saved XLSX archive. | Extend the example to automatically remove any .bin entries found in the XLSX package after saving.
// Common Searches: Aspose.Cells how to check XLSX package for unwanted .bin files in C# | C# verify that saved Excel file does not include printer settings binaries | detect and delete .bin entries inside an XLSX archive using Aspose.Cells | ensure clean workbook package without printer‑settings files after saving with Aspose.Cells | scan ZIP contents of generated XLSX for binary files in .NET
// Tags: Aspose.Cells save workbook to XLSX and inspect archive | C# scan XLSX ZIP entries for .bin files | remove printer settings binaries from Excel package | validate Excel workbook package integrity with Aspose.Cells | detect unwanted binary files in generated XLSX

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// Loads (or creates) a workbook, saves it as a new XLSX file using Aspose.Cells, opens the file as a ZipArchive, enumerates its entries to find any .bin files (printer‑settings), and outputs whether such files are present.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string inputPath = "cleaned_workbook.xlsx";

            // Ensure the input file exists; create a blank workbook if it does not.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}. Creating a new empty workbook.");
                Workbook newWb = new Workbook();
                newWb.Save(inputPath, SaveFormat.Xlsx);
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as a new XLSX file
            string outputPath = "cleaned_workbook_saved.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Scan the resulting package for any .bin files (printer settings)
            bool binFileFound = false;
            using (FileStream fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
                    {
                        binFileFound = true;
                        Console.WriteLine($"Unexpected binary file found: {entry.FullName}");
                    }
                }
            }

            Console.WriteLine(binFileFound
                ? "Printer‑settings .bin files exist in the XLSX package."
                : "No printer‑settings .bin files were found in the XLSX package.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
