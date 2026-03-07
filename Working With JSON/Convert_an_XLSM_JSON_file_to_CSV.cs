using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source JSON file (exported from an XLSM workbook)
        string sourcePath = "input.json";

        // If the JSON file does not exist, create a simple workbook and export it to JSON
        if (!File.Exists(sourcePath))
        {
            Workbook tempWb = new Workbook();
            tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
            tempWb.Worksheets[0].Cells["B1"].PutValue(123);
            tempWb.Save(sourcePath, SaveFormat.Json);
        }

        // Path where the CSV file will be saved
        string csvPath = "output.csv";

        // Load the workbook from JSON representation
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Save the workbook as CSV
        workbook.Save(csvPath, SaveFormat.Csv);

        Console.WriteLine($"Conversion completed successfully. CSV saved to: {csvPath}");
    }
}