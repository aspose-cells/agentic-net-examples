using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonToExcelConverter
{
    static void Main(string[] args)
    {
        // Input JSON file (plain text) and output Excel file paths
        string inputJsonPath = "input.json";   // replace with your JSON file path
        string outputExcelPath = "output.xlsx"; // replace with desired Excel file path

        // Read the entire JSON content from the text file
        string jsonContent = File.ReadAllText(inputJsonPath);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure layout options: treat JSON arrays as tables
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, layoutOptions);

        // Save the workbook to the specified Excel file
        workbook.Save(outputExcelPath);
    }
}