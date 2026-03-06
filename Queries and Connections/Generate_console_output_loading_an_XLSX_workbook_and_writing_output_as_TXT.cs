using System;
using Aspose.Cells;

namespace AsposeCellsTxtExportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path for the output TXT file
            string outputPath = "output.txt";

            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(sourcePath);

            // Display basic information about the loaded workbook
            Console.WriteLine($"Workbook loaded. Worksheets count: {workbook.Worksheets.Count}");

            // Create TXT save options (using CSV format for text output)
            TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
            // Export all sheets to the text file (set to true if you want all sheets)
            txtOptions.ExportAllSheets = true;

            // Save the workbook as a TXT (CSV) file using the specified options
            workbook.Save(outputPath, txtOptions);

            Console.WriteLine($"Workbook saved as TXT to: {outputPath}");
        }
    }
}