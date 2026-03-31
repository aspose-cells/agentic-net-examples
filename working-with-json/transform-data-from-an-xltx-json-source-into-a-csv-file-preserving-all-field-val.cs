using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the JSON file that contains the Excel‑like data
        string jsonFilePath = "source.json";

        // Read the entire JSON content
        string jsonContent = File.ReadAllText(jsonFilePath);

        // Create a new empty workbook
        Workbook workbook = new Workbook();

        // Prepare JSON layout options – treat arrays as tables to keep all fields
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        // Import the JSON data into the first worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, workbook.Worksheets[0].Cells, 0, 0, layoutOptions);

        // Define the output CSV file path
        string csvOutputPath = "output.csv";

        // Save the workbook as CSV, preserving all field values
        workbook.Save(csvOutputPath, SaveFormat.Csv);

        Console.WriteLine($"JSON data has been converted to CSV at: {csvOutputPath}");
    }
}