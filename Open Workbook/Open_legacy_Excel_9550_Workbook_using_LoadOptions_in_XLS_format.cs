using System;
using Aspose.Cells;

namespace AsposeCellsLegacyExample
{
    class Program
    {
        static void Main()
        {
            // Path to the legacy Excel 95/5.0 workbook (XLS format)
            string legacyFilePath = "legacy.xls";

            // Create LoadOptions specifying the Excel 97-2003 (XLS) format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(legacyFilePath, loadOptions);

            // Example operation: read the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook back to XLS using XlsSaveOptions
            XlsSaveOptions saveOptions = new XlsSaveOptions();
            workbook.Save("converted.xls", saveOptions);

            Console.WriteLine("Legacy workbook loaded and saved successfully.");
        }
    }
}