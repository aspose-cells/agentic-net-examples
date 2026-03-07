using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Input JSON file (generated from MHTML) and output Excel file paths
        string jsonFilePath = "input.json";
        string excelFilePath = "output.xlsx";

        // Read the entire JSON content from the file
        string jsonContent = File.ReadAllText(jsonFilePath);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure JSON import options
        JsonLayoutOptions importOptions = new JsonLayoutOptions();
        importOptions.ArrayAsTable = true;      // Treat JSON arrays as tables
        importOptions.IgnoreTitle = true;       // Ignore the root title if present (optional)

        // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, importOptions);

        // Save the workbook as an Excel file (XLSX format)
        workbook.Save(excelFilePath, SaveFormat.Xlsx);
    }
}