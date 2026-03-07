using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source JSON file (exported from an XLTX workbook)
        string jsonPath = "source.json";

        // Desired CSV output path
        string csvPath = "result.csv";

        // Ensure the JSON file exists; if not, create a simple sample JSON
        if (!File.Exists(jsonPath))
        {
            string sampleJson = @"{
                ""Sheet1"": [
                    { ""Name"": ""Alice"", ""Age"": 30 },
                    { ""Name"": ""Bob"",   ""Age"": 25 }
                ]
            }";
            File.WriteAllText(jsonPath, sampleJson);
        }

        // Read the entire JSON content from the file
        string jsonContent = File.ReadAllText(jsonPath);

        // Create a new empty workbook
        Workbook workbook = new Workbook();

        // Configure JSON import options: treat JSON arrays as tables
        JsonLayoutOptions importOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        // Import the JSON data into the first worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, workbook.Worksheets[0].Cells, 0, 0, importOptions);

        // Save the populated workbook as CSV
        workbook.Save(csvPath, SaveFormat.Csv);
    }
}