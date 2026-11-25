// Title: Disable text wrap for all chart data labels in every worksheet of an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# disable data label text wrapping for all charts in a workbook | How to turn off text wrap on chart data labels across multiple sheets with Aspose.Cells | Batch update Excel charts to set DataLabels.IsTextWrapped = false using .NET | Iterate worksheets and charts to remove data label wrapping in Aspose.Cells | C# code to process an Excel file and disable chart data label wrapping globally
// Tags: Aspose.Cells chart data label wrap removal | DataLabels.IsTextWrapped false .NET | batch update Excel chart settings Aspose.Cells | iterate worksheets charts Aspose.Cells | automate chart label formatting C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartDataLabelWrapDisabler
{
    // The example loads an Excel workbook, iterates through each worksheet and every chart it contains, sets DataLabels.IsTextWrapped to false for all series to turn off text wrapping, and saves the modified workbook as a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath;
                string outputPath;

                if (args.Length >= 1)
                {
                    inputPath = args[0];
                    Console.WriteLine($"Input workbook: {inputPath}");
                }
                else
                {
                    Console.Write("Enter the full path of the Excel workbook to process: ");
                    inputPath = Console.ReadLine()?.Trim('"') ?? string.Empty;
                }

                if (string.IsNullOrWhiteSpace(inputPath) || !File.Exists(inputPath))
                {
                    Console.Error.WriteLine("Error: The specified input file does not exist.");
                    return;
                }

                // Determine output path
                if (args.Length >= 2)
                {
                    outputPath = args[1];
                }
                else
                {
                    string directory = Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();
                    string filenameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                    string ext = Path.GetExtension(inputPath);
                    outputPath = Path.Combine(directory, $"{filenameWithoutExt}_NoWrap{ext}");
                }

                Console.WriteLine($"Output workbook will be saved as: {outputPath}");

                ProcessWorkbook(inputPath, outputPath);
                Console.WriteLine("Processing completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="inputPath">Path to the source workbook.</param>
        /// <param name="outputPath">Path where the modified workbook will be saved.</param>
        private static void ProcessWorkbook(string inputPath, string outputPath)
        {
            try
            {
                // Load the workbook (file existence already verified)
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Console.WriteLine($"Processing worksheet: {sheet.Name}");

                    // Iterate through each chart in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        Console.WriteLine($"  Processing chart of type {chart.Type}");

                        // Iterate through each series in the chart
                        foreach (Series series in chart.NSeries)
                        {
                            // Access the DataLabels object (always non‑null)
                            DataLabels dataLabels = series.DataLabels;

                            // Disable text wrapping
                            dataLabels.IsTextWrapped = false;

                            Console.WriteLine($"    DataLabels.IsTextWrapped set to false for series.");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error processing workbook: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
                throw; // Re‑throw to allow outer handler to report
            }
        }
    }
}
