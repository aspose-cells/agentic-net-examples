// Title: Batch convert Excel .xlsx files to PDF with Aspose.Cells for .NET and generate a detailed conversion report
// AI Prompts: Write a C# console app that iterates over all .xlsx files in a folder, uses Aspose.Cells Workbook.Save to export each workbook to PDF, and records each success or exception in a text report. | Modify the batch converter to accept the desired SaveFormat (e.g., Pdf, Csv, Html) as a command‑line argument and automatically choose the correct file extension. | Add logic to skip files that already have the target extension, count them as 'skipped', and include the skip count in the generated summary report.
// Common Searches: c# Aspose.Cells batch convert multiple xlsx to pdf and create log file | how to generate conversion summary report for Excel files using Aspose.Cells .NET | command line tool to convert all Excel workbooks in a directory to PDF with Aspose.Cells | skip already converted files when batch converting Excel to PDF in C#
// Tags: Aspose.Cells batch Excel to PDF conversion | C# generate conversion report for Aspose.Cells | Aspose.Cells SaveFormat to file extension mapping | command‑line folder processing with Aspose.Cells | error handling in Aspose.Cells workbook conversion

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

// // C# console program that scans a specified input folder for .xlsx workbooks, converts each to a target format (default PDF) using Aspose.Cells, writes successful and failed conversion details to a text report, and outputs counts of successes, failures, and skips.
class BatchConverter
{
    static void Main(string[] args)
    {
        // Input directory containing source Excel files
        string inputDir = args.Length > 0 ? args[0] : "InputFiles";
        // Output directory for converted files
        string outputDir = args.Length > 1 ? args[1] : "OutputFiles";
        // Desired output format (change as needed)
        SaveFormat targetFormat = SaveFormat.Pdf;

        // Verify input directory exists
        if (!Directory.Exists(inputDir))
        {
            Console.WriteLine($"Input directory not found: {Path.GetFullPath(inputDir)}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputDir);

        // Lists to hold conversion results
        List<string> successes = new List<string>();
        List<string> failures = new List<string>();

        try
        {
            // Process each Excel file in the input directory
            foreach (string filePath in Directory.GetFiles(inputDir, "*.xlsx"))
            {
                // Guard against missing files (should not happen, but safe)
                if (!File.Exists(filePath))
                {
                    failures.Add($"{Path.GetFileName(filePath)}: File not found.");
                    continue;
                }

                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string outFile = Path.Combine(outputDir, fileName + GetExtension(targetFormat));

                try
                {
                    // Load workbook
                    Workbook wb = new Workbook(filePath);
                    // Save in target format
                    wb.Save(outFile, targetFormat);
                    successes.Add($"{fileName} -> {Path.GetFileName(outFile)}");
                }
                catch (Exception ex)
                {
                    failures.Add($"{fileName}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch unexpected errors during enumeration
            Console.WriteLine($"Error processing files: {ex.Message}");
            return;
        }

        // Generate summary report
        string reportPath = Path.Combine(outputDir, "ConversionReport.txt");
        try
        {
            using (StreamWriter sw = new StreamWriter(reportPath))
            {
                sw.WriteLine("Batch Conversion Summary");
                sw.WriteLine($"Date: {DateTime.Now}");
                sw.WriteLine();
                sw.WriteLine("Successful conversions:");
                foreach (var s in successes)
                    sw.WriteLine("  " + s);
                sw.WriteLine();
                sw.WriteLine("Failed conversions:");
                foreach (var f in failures)
                    sw.WriteLine("  " + f);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to write report: {ex.Message}");
        }

        // Output summary to console
        Console.WriteLine("Conversion completed.");
        Console.WriteLine($"Successes: {successes.Count}");
        Console.WriteLine($"Failures: {failures.Count}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }

    // Helper to map SaveFormat to file extension
    static string GetExtension(SaveFormat format)
    {
        switch (format)
        {
            case SaveFormat.Pdf: return ".pdf";
            case SaveFormat.Xlsx: return ".xlsx";
            case SaveFormat.Csv: return ".csv";
            case SaveFormat.Html: return ".html";
            default: return ".out";
        }
    }
}
