using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonToCsv
{
    class Program
    {
        static void Main()
        {
            // Paths for source JSON file and destination CSV file
            string jsonFilePath = "input.json";
            string csvFilePath = "output.csv";

            // Load the JSON file into a Workbook using JsonLoadOptions
            JsonLoadOptions loadOptions = new JsonLoadOptions();
            Workbook workbook = new Workbook(jsonFilePath, loadOptions);

            // Save the workbook content as CSV
            workbook.Save(csvFilePath, SaveFormat.Csv);

            Console.WriteLine($"JSON file '{jsonFilePath}' has been converted to CSV at '{csvFilePath}'.");
        }
    }
}