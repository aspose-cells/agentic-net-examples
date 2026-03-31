using System;
using Aspose.Cells;

namespace AsposeCellsCsvImportDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be imported
            string csvPath = "data.csv";

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Import CSV data into the worksheet starting at cell A1 (row 0, column 0)
            // Parameters: file name, delimiter, convert numeric strings, first row, first column
            worksheet.Cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Save the workbook as an XLSX file
            workbook.Save("ImportedFromCsv.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("CSV data has been imported and saved to 'ImportedFromCsv.xlsx'.");
        }
    }
}