using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the DBF‑formatted CSV file
        string csvPath = "data.dbf.csv";

        // Ensure the CSV file exists; create a simple sample if it does not
        if (!File.Exists(csvPath))
        {
            File.WriteAllText(csvPath,
                "Id,Name,Value\n" +
                "1,Alpha,100\n" +
                "2,Beta,200\n" +
                "3,Gamma,300");
        }

        // Create a new workbook and get the first worksheet's cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the CSV data (comma delimiter, convert numeric values)
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Determine the used range after import
        var usedRange = cells.MaxDisplayRange;

        // Configure JSON export options
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportEmptyCells = true,   // include empty cells in JSON
            HasHeaderRow = true,       // first row contains column names
            ToExcelStruct = false      // regular JSON object array
        };

        // Export the used range to a JSON string
        string json = JsonUtility.ExportRangeToJson(usedRange, jsonOptions);

        // Write the JSON string to a file
        File.WriteAllText("output.json", json);

        // Optional: save the workbook as Excel for verification
        workbook.Save("intermediate.xlsx", SaveFormat.Xlsx);
    }
}