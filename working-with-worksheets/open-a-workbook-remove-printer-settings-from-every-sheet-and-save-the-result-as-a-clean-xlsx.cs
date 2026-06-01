using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemovePrinterSettingsDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_clean.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and clear the printer settings
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // The PrinterSettings property holds printer configuration as a byte array.
                    // Setting it to null removes any stored printer settings for the sheet.
                    sheet.PageSetup.PrinterSettings = null;
                }

                // Save the cleaned workbook as a regular XLSX file
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Printer settings removed and workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}