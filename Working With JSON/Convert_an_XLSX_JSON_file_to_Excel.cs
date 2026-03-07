using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source JSON file (exported from an Excel workbook)
        string jsonFilePath = "source.json";

        // Desired output Excel file path
        string excelFilePath = "converted.xlsx";

        // Read the entire JSON content
        string jsonContent = File.ReadAllText(jsonFilePath);

        // Create a new, empty workbook
        Workbook workbook = new Workbook();

        // Get the Cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Configure import options – treat JSON arrays as tables
        JsonLayoutOptions importOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        // Import the JSON data starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, cells, 0, 0, importOptions);

        // Save the workbook in XLSX format
        workbook.Save(excelFilePath, SaveFormat.Xlsx);

        Console.WriteLine($"JSON file has been successfully converted to Excel: {excelFilePath}");
    }
}