using System;
using Aspose.Cells;

namespace AsposeCellsCsvLoadDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be loaded
            string csvPath = "input.csv";

            // Create load options specifying that the source format is CSV
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

            // Initialize a workbook by loading the CSV file with the specified options
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Access the first worksheet to verify that data was loaded
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("First cell value: " + sheet.Cells["A1"].StringValue);
            Console.WriteLine("Second cell value: " + sheet.Cells["B1"].StringValue);

            // Save the workbook as an XLSX file (conversion from CSV to Excel)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}