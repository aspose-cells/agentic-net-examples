using System;
using Aspose.Cells;

namespace AsposeCellsLoadXlsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel 97‑2003 workbook (XLS file)
            string filePath = "sample.xls";

            // Initialize LoadOptions with the specific format for Excel 97‑2003 files
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: read the value of cell A1 and display it
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);
        }
    }
}