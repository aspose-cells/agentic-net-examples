using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "sample.xlsx";

            // Create LoadOptions for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Specify the desired culture (e.g., German - uses comma as decimal separator)
            loadOptions.CultureInfo = new CultureInfo("de-DE");

            // Load the workbook with the custom culture settings
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Access a cell that contains a numeric value
            Cell cell = workbook.Worksheets[0].Cells["A1"];

            // Display the value as interpreted with the German culture
            Console.WriteLine("Value with German culture: " + cell.StringValue);

            // Save the workbook (culture setting affects only the loaded data, not the saved file)
            string outputPath = "sample_loaded_de-DE.xlsx";
            workbook.Save(outputPath);
        }
    }
}