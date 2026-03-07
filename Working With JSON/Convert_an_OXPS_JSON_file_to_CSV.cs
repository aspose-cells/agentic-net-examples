using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file (generated from OXPS)
            string sourceJsonPath = "input.json";

            // Path for the resulting CSV file
            string destCsvPath = "output.csv";

            try
            {
                // Load the JSON file into a workbook using JsonLoadOptions
                JsonLoadOptions loadOptions = new JsonLoadOptions();
                Workbook workbook = new Workbook(sourceJsonPath, loadOptions);

                // Save the workbook as CSV
                workbook.Save(destCsvPath, SaveFormat.Csv);

                Console.WriteLine($"Conversion completed: '{sourceJsonPath}' -> '{destCsvPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}