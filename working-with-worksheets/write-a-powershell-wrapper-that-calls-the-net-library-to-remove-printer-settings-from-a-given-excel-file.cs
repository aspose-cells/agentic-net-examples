using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPrinterCleaner
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input file path and output file path
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsPrinterCleaner <inputFile> <outputFile>");
                return;
            }

            string inputFile = args[0];
            string outputFile = args[1];

            // Verify that the input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook from the input file
                workbook = new Workbook(inputFile);

                // Clear printer settings for each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // PageSetup.PrinterSettings is a byte[]; assign an empty array to remove settings
                    sheet.PageSetup.PrinterSettings = new byte[0];
                }

                // Save the modified workbook to the output file
                workbook.Save(outputFile);
                Console.WriteLine($"Printer settings removed successfully. Output saved to: {outputFile}");
            }
            catch (Exception ex)
            {
                // Report any runtime errors
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                workbook?.Dispose();
            }
        }
    }
}