using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsOpenExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the modern Excel workbook (XLSX or XLSB)
            string filePath = @"C:\Data\SampleWorkbook.xlsx";

            // Ensure the file exists; if not, create a simple workbook and save it.
            if (!File.Exists(filePath))
            {
                // Create directory if it doesn't exist
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));

                // Create a new workbook with a default worksheet
                var tempWorkbook = new Workbook();
                tempWorkbook.Worksheets[0].Name = "Sheet1";
                tempWorkbook.Save(filePath, SaveFormat.Xlsx);
            }

            // Load the workbook with explicit XLSX format
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(filePath, loadOptions);

            // Example usage: display the number of worksheets loaded
            Console.WriteLine($"Workbook loaded successfully. Worksheets count: {workbook.Worksheets.Count}");
        }
    }
}