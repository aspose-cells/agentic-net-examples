using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure JSON layout options to treat arrays as tables (columns = properties)
        JsonLayoutOptions jsonOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true,          // map each object property to a column
            ConvertNumericOrDate = true   // optional: convert numbers and dates automatically
        };

        // Read the JSON content from a file
        string jsonFilePath = "input.json";
        string jsonContent = File.ReadAllText(jsonFilePath);

        // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, jsonOptions);

        // Save the workbook as a CSV file
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}