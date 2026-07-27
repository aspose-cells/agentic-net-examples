// Title: C# Batch Convert Excel (.xlsx) to PDF with Per‑File Progress Using Aspose.Cells
// Description: A console program that loops through parallel source and destination arrays, uses Aspose.Cells ConversionUtility to convert each workbook to PDF, calculates and displays start/end percentages for every file, and logs errors without stopping the batch.
// Keywords: Aspose.Cells batch conversion C# | Excel to PDF conversion progress | ConversionUtility multiple files | C# per‑file percentage reporting | resilient Excel PDF export
// Common Searches: C# batch convert .xlsx to .pdf with progress Aspose.Cells | How to show percentage while converting Excel files to PDF in C# | Aspose.Cells ConversionUtility example for multiple files | Error‑tolerant Excel to PDF batch conversion C#
// Developer Intent: Convert a collection of Excel workbooks to PDF while reporting the conversion percentage for each file and handling errors gracefully.
// Use Cases: Automate nightly report generation by converting several Excel files to PDFs and displaying progress in a console log. | Add a fault‑tolerant conversion step to an ETL pipeline that continues processing remaining files after a failure. | Provide a lightweight utility for end‑users to monitor batch conversion status with clear start/end percentages.
// AI Prompts: Generate C# code that uses Aspose.Cells ConversionUtility to convert an array of .xlsx files to .pdf and prints start and end percentages for each file. | Create a robust batch conversion routine that logs conversion errors and proceeds with remaining files using Aspose.Cells. | Write a method that accepts source and destination file lists, performs conversions with Aspose.Cells, and returns a summary of successes and failures.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace BatchConversionWithProgress
{
    // Simple program that converts a list of Excel files to PDF format
    // and reports the percentage completed for each file.
    // A console program that loops through parallel source and destination arrays, uses Aspose.Cells ConversionUtility to convert each workbook to PDF, calculates and displays start/end percentages for every file, and logs errors without stopping the batch.
    class Program
    {
        static void Main(string[] args)
        {
            // Define source Excel files (ensure these files exist)
            string[] sourceFiles = new string[]
            {
                "Report1.xlsx",
                "Report2.xlsx",
                "Report3.xlsx"
            };

            // Define corresponding destination files (PDF format)
            string[] destFiles = new string[]
            {
                "Report1.pdf",
                "Report2.pdf",
                "Report3.pdf"
            };

            // Validate that source and destination arrays have the same length
            if (sourceFiles.Length != destFiles.Length)
            {
                Console.WriteLine("Source and destination file arrays must have the same number of elements.");
                return;
            }

            int totalFiles = sourceFiles.Length;

            for (int i = 0; i < totalFiles; i++)
            {
                // Calculate and display progress before starting the conversion
                double startPercent = (i * 100.0) / totalFiles;
                Console.WriteLine($"Starting conversion {i + 1}/{totalFiles} ({startPercent:F2}% completed).");
                Console.WriteLine($"Source: {sourceFiles[i]}");
                Console.WriteLine($"Destination: {destFiles[i]}");

                try
                {
                    // Perform the conversion using Aspose.Cells ConversionUtility
                    ConversionUtility.Convert(sourceFiles[i], destFiles[i]);

                    // Report completion for this file
                    double endPercent = ((i + 1) * 100.0) / totalFiles;
                    Console.WriteLine($"Finished conversion {i + 1}/{totalFiles} ({endPercent:F2}% completed).");
                }
                catch (Exception ex)
                {
                    // Report any errors but continue processing remaining files
                    Console.WriteLine($"Error converting file '{sourceFiles[i]}': {ex.Message}");
                }

                Console.WriteLine(new string('-', 50));
            }

            Console.WriteLine("Batch conversion process completed.");
        }
    }
}
