using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents an Excel workbook
            string sourceJsonPath = "input.json";

            // Path for the resulting CSV file
            string destCsvPath = "output.csv";

            // Load the JSON file into a Workbook using default JsonLoadOptions
            JsonLoadOptions loadOptions = new JsonLoadOptions(); // default options
            Workbook workbook = new Workbook(sourceJsonPath, loadOptions);

            // Save the workbook as CSV with default options
            workbook.Save(destCsvPath, SaveFormat.Csv);

            Console.WriteLine($"Conversion completed: {sourceJsonPath} -> {destCsvPath}");
        }
    }
}