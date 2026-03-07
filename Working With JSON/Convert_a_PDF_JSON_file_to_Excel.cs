using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Paths for the source JSON file (exported from PDF) and the target Excel file
        string jsonFilePath = "source.json";
        string excelFilePath = "result.xlsx";

        // Read the entire JSON content from the file
        string jsonContent = File.ReadAllText(jsonFilePath);

        // Create a new workbook (Aspose.Cells) and obtain the Cells collection of the first worksheet
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Configure JSON layout options – treat JSON arrays as tables for proper column mapping
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true
        };

        // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, cells, 0, 0, layoutOptions);

        // Save the populated workbook as an Excel file (XLSX)
        workbook.Save(excelFilePath);
    }
}