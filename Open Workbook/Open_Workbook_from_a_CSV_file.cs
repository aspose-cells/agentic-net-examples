using System;
using Aspose.Cells;

namespace AsposeCellsCsvExample
{
    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be opened
            string csvFilePath = "input.csv";

            // Open the CSV file as a workbook using the string constructor
            Workbook workbook = new Workbook(csvFilePath);

            // Access the first worksheet (the CSV data is loaded into the first sheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: display the value of the first cell (A1) to verify the load
            Console.WriteLine("A1 value: " + worksheet.Cells["A1"].StringValue);

            // Save the workbook in XLSX format (optional, demonstrates the save lifecycle)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}