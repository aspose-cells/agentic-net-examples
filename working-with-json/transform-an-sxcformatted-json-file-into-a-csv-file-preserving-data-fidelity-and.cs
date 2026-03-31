using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonToCsvConverter
{
    static void Main()
    {
        // Input JSON file and output CSV file paths
        string jsonPath = "input.json";
        string csvPath = "output.csv";

        // Read the entire JSON content
        string jsonContent = File.ReadAllText(jsonPath);

        // Create a new workbook (in-memory Excel file)
        Workbook workbook = new Workbook();

        // Get the Cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Configure JSON import options: treat a JSON array as a table
        JsonLayoutOptions importOptions = new JsonLayoutOptions();
        importOptions.ArrayAsTable = true;

        // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, cells, 0, 0, importOptions);

        // Save the workbook as CSV, preserving column mapping and data fidelity
        workbook.Save(csvPath, SaveFormat.Csv);

        Console.WriteLine($"Conversion completed: '{jsonPath}' → '{csvPath}'");
    }
}