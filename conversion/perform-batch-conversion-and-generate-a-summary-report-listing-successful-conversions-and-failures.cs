// Title: C# batch conversion of multiple Excel (.xlsx) files to PDF with Aspose.Cells and generate a detailed conversion report
// Common Searches: C# code to convert all .xlsx files in a folder to PDF using Aspose.Cells | how to create a conversion log for batch Excel to PDF conversion in .NET | Aspose.Cells batch processing example with error handling and report file | save multiple workbooks as PDF and write status to a text file using Aspose.Cells | automate Excel to PDF conversion for many files and capture failures in C#
// Tags: batch excel to pdf conversion Aspose.Cells | c# generate conversion report for workbook saves | aspocells saveformat pdf batch operation | c# exception handling for workbook conversion | process multiple xlsx files with Aspose.Cells

using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Cells;

namespace BatchConversion
{
    // The program scans a specified input directory for .xlsx files, loads each workbook with Aspose.Cells, saves it as a PDF in an output folder, and writes a tab‑delimited report that records the file name, conversion status (Success or Failed), and any error messages.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing source Excel files
            string inputFolder = @"C:\Input";

            // Output folder where converted files and the report will be saved
            string outputFolder = @"C:\Output";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Path for the summary report
            string reportPath = Path.Combine(outputFolder, "ConversionReport.txt");

            // Collection to hold report lines
            List<string> reportLines = new List<string>();

            // Header for the report
            reportLines.Add("FileName\tStatus\tMessage");

            // Process each Excel file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                string fileName = Path.GetFileName(filePath);
                try
                {
                    // Load the workbook (load rule)
                    Workbook workbook = new Workbook(filePath);

                    // Determine the output file path (PDF in this example)
                    string outputFilePath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                    // Save the workbook in the desired format (save rule)
                    workbook.Save(outputFilePath, SaveFormat.Pdf);

                    // Record successful conversion
                    reportLines.Add($"{fileName}\tSuccess\t");
                }
                catch (Exception ex)
                {
                    // Record failure with the exception message
                    reportLines.Add($"{fileName}\tFailed\t{ex.Message}");
                }
            }

            // Write the summary report to disk
            File.WriteAllLines(reportPath, reportLines);
        }
    }
}
