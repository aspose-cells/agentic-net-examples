using System;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook (XLSX)
            string sourcePath = "input.xlsx";

            // Desired output CSV file path
            string outputPath = "output.csv";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Configure CSV save options to trim leading blank rows and columns
            TxtSaveOptions csvOptions = new TxtSaveOptions
            {
                TrimLeadingBlankRowAndColumn = true
            };

            // Save the workbook as CSV using the configured options
            workbook.Save(outputPath, csvOptions);

            Console.WriteLine("Conversion completed: " + outputPath);
        }
    }
}