using System;
using Aspose.Cells;

namespace AsposeCellsConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the resulting CSV file
            string destinationPath = "output.csv";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Create CSV save options and enable separators for blank rows
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                KeepSeparatorsForBlankRow = true
            };

            // Save the workbook as CSV using the configured options
            workbook.Save(destinationPath, csvOptions);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destinationPath}'");
        }
    }
}