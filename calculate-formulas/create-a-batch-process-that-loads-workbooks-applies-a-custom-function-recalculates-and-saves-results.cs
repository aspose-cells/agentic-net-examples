// Title: C# batch processing of multiple Excel workbooks: load, add formula, recalculate, and save with Aspose.Cells
// Common Searches: C# Aspose.Cells batch process all .xlsx files in a folder | how to add a formula to each workbook and recalculate using Aspose.Cells | automate loading, calculating, and saving multiple Excel workbooks with Aspose.Cells .NET | sample code for batch applying formulas to Excel files in C#
// Tags: batch load Excel workbooks Aspose.Cells | apply addition formula to cells C# | recalculate all workbook formulas Aspose.Cells | save processed Excel files to output directory C# | process multiple .xlsx files Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcess
{
    // The example iterates through every .xlsx file in a specified input folder, loads each workbook with Aspose.Cells, inserts sample values, sets an addition formula in cell A1, recalculates all formulas, and saves the modified workbook to an output folder, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Folder containing source workbooks.
                string inputFolder = @"C:\InputWorkbooks";
                // Folder where processed workbooks will be saved.
                string outputFolder = @"C:\OutputWorkbooks";

                // Ensure input and output directories exist.
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }
                Directory.CreateDirectory(outputFolder);

                // Get all Excel files in the input folder.
                string[] files = Directory.GetFiles(inputFolder, "*.xlsx");

                foreach (string filePath in files)
                {
                    try
                    {
                        // Verify the file exists before loading.
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Load the workbook.
                        Workbook workbook = new Workbook(filePath);

                        // Example: apply a simple addition formula to cell A1 in the first worksheet.
                        Worksheet sheet = workbook.Worksheets[0];
                        // Set up sample values in B1 and C1.
                        sheet.Cells["B1"].PutValue(10);
                        sheet.Cells["C1"].PutValue(20);
                        // Use built‑in addition formula in A1: =B1+C1
                        sheet.Cells["A1"].Formula = "B1+C1";

                        // Recalculate all formulas in the workbook.
                        workbook.CalculateFormula();

                        // Build output file path.
                        string fileName = Path.GetFileName(filePath);
                        string outputPath = Path.Combine(outputFolder, fileName);

                        // Save the modified workbook.
                        workbook.Save(outputPath);
                        Console.WriteLine($"Processed and saved: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
