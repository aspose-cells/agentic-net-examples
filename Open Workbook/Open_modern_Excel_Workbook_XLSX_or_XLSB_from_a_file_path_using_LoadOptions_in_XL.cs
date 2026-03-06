using System;
using Aspose.Cells;

namespace AsposeCellsOpenExample
{
    class Program
    {
        static void Main()
        {
            // Path to the modern Excel file (XLSX or XLSB)
            string filePath = "sample.xlsx";

            // Create LoadOptions specifying the XLSX format
            // This uses the LoadOptions(LoadFormat) constructor rule
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Open the workbook with the file path and the load options
            // This uses the Workbook(string, LoadOptions) constructor rule
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Demonstrate that the workbook is loaded
            Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
        }
    }
}