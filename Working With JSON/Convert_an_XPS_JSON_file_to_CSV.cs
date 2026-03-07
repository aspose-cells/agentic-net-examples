using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class XpsJsonToCsvConverter
{
    static void Main()
    {
        // Paths for the source JSON file (generated from XPS) and the target CSV file
        string jsonFilePath = "input.json";
        string csvFilePath = "output.csv";

        // Read the entire JSON content from the file
        string jsonContent = File.ReadAllText(jsonFilePath);

        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure JSON import options: treat JSON arrays as tables
        JsonLayoutOptions importOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, cells, 0, 0, importOptions);

        // Save the populated workbook as CSV
        workbook.Save(csvFilePath, SaveFormat.Csv);

        Console.WriteLine("Conversion from XPS JSON to CSV completed successfully.");
    }
}