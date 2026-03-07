using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents an XLSM workbook
            string sourceJsonPath = "input.json";

            // Desired CSV output path
            string csvOutputPath = "output.csv";

            // Load the JSON file using JsonLoadOptions (default options)
            JsonLoadOptions loadOptions = new JsonLoadOptions();

            // Create a Workbook instance from the JSON file
            Workbook workbook = new Workbook(sourceJsonPath, loadOptions);

            // Save the workbook as CSV using the default CSV format
            workbook.Save(csvOutputPath, SaveFormat.Csv);

            Console.WriteLine($"Conversion completed: '{sourceJsonPath}' → '{csvOutputPath}'");
        }
    }
}