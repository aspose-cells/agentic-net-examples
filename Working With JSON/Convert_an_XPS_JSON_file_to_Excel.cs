using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

public class XpsJsonToExcelConverter
{
    // Converts a JSON file (originating from XPS) to an Excel workbook.
    public static void Convert(string jsonFilePath, string excelOutputPath)
    {
        // Read the JSON content from the source file
        string jsonContent = File.ReadAllText(jsonFilePath);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Configure JSON import options
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            // Import JSON arrays as tables (each object becomes a row)
            ArrayAsTable = true
        };

        // Import the JSON data starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, sheet.Cells, 0, 0, layoutOptions);

        // Save the workbook as an XLSX file
        workbook.Save(excelOutputPath, SaveFormat.Xlsx);
    }

    // Example usage
    public static void Main()
    {
        string sourceJson = "input.json";          // Path to the XPS‑derived JSON file
        string destinationExcel = "output.xlsx";   // Desired Excel output path

        Convert(sourceJson, destinationExcel);
        Console.WriteLine("Conversion completed successfully.");
    }
}