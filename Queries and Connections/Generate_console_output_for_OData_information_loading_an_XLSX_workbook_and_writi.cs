using System;
using System.IO;
using Aspose.Cells;

namespace ODataWorkbookToTxt
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Verify the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook using the constructor that accepts a file path
            Workbook workbook = new Workbook(sourcePath);

            // Output basic workbook information to the console
            Console.WriteLine($"Workbook loaded from: {sourcePath}");
            Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
            Console.WriteLine($"Active worksheet: {workbook.Worksheets.ActiveSheetIndex + 1} - {workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex].Name}");

            // Example of reading a cell value (A1) from the active sheet
            var activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];
            var cellValue = activeSheet.Cells["A1"].StringValue;
            Console.WriteLine($"Value of cell A1 in active sheet: {cellValue}");

            // Prepare TxtSaveOptions to export the workbook as a text (CSV) file
            TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Export all sheets; set to false to export only the active sheet
                ExportAllSheets = true
            };

            // Path for the output text file
            string outputPath = "output.txt";

            // Save the workbook to the text file using the provided Save method overload
            workbook.Save(outputPath, txtOptions);

            Console.WriteLine($"Workbook successfully exported to text file: {outputPath}");
        }
    }
}