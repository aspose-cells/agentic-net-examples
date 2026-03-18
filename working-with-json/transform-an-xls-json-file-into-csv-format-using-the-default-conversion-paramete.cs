using System;
using Aspose.Cells;

namespace AsposeCellsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source JSON file that represents an Excel workbook
            string sourceJson = "input.json";

            // Desired CSV output file path
            string outputCsv = "output.csv";

            // Load the JSON file into a Workbook using default JsonLoadOptions
            Workbook workbook = new Workbook(sourceJson, new JsonLoadOptions());

            // Save the workbook as CSV using the default conversion parameters
            workbook.Save(outputCsv, SaveFormat.Csv);

            Console.WriteLine($"Conversion completed: '{sourceJson}' -> '{outputCsv}'");
        }
    }
}