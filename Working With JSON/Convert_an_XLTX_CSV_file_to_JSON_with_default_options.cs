using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExamples
{
    public class CsvToJsonConversion
    {
        public static void Main()
        {
            // Path to the source CSV file
            string sourceCsvPath = "input.csv";

            // Path where the resulting JSON file will be saved
            string outputJsonPath = "output.json";

            // Load the CSV file into a workbook
            var loadOptions = new LoadOptions(LoadFormat.CSV);
            var workbook = new Workbook(sourceCsvPath, loadOptions);

            // Save the workbook as JSON
            var saveOptions = new JsonSaveOptions();
            workbook.Save(outputJsonPath, saveOptions);

            Console.WriteLine($"Conversion completed. JSON file saved to: {outputJsonPath}");
        }
    }
}