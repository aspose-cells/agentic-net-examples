using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureLoadDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the XLSX file to be loaded
            string filePath = "input.xlsx";

            // Create LoadOptions for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Set the desired culture (e.g., German uses comma as decimal separator)
            loadOptions.CultureInfo = new CultureInfo("de-DE");

            // Load the workbook with the specified culture settings
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access a cell to demonstrate that values are parsed using the German culture
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];

            // Output the raw value and its string representation
            Console.WriteLine($"Cell A1 raw value: {cell.Value}");
            Console.WriteLine($"Cell A1 string value (culture-aware): {cell.StringValue}");
        }
    }
}