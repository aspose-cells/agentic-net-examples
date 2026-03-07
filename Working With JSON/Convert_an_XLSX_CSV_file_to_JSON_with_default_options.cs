using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string sourceCsv = "input.csv";
        // Desired path for the output JSON file
        string outputJson = "output.json";

        // Create a sample CSV file (optional, for demonstration)
        System.IO.File.WriteAllText(sourceCsv, "Name,Age\nJohn,30\nAlice,25");

        // Load options specifying that the source file is a CSV
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

        // Save options for JSON with default settings
        JsonSaveOptions saveOptions = new JsonSaveOptions();

        // Convert the CSV file to JSON using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourceCsv, loadOptions, outputJson, saveOptions);

        Console.WriteLine($"Conversion completed: {sourceCsv} -> {outputJson}");
    }
}