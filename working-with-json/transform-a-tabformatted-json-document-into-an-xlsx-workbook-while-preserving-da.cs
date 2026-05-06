using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonToExcel
{
    static void Main()
    {
        // Input TAB‑formatted JSON file path
        string jsonPath = "input.json";

        // Output XLSX workbook path
        string excelPath = "output.xlsx";

        // Read the JSON content (tabs are retained as whitespace)
        string jsonContent = File.ReadAllText(jsonPath);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Configure layout options to preserve data types and treat arrays as tables
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true,            // Process JSON arrays as tables
            ConvertNumericOrDate = true,    // Convert numeric and date strings to proper types
            NumberFormat = "0.##",          // Optional numeric format
            DateFormat = "yyyy-MM-dd"       // Optional date format
        };

        // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, sheet.Cells, 0, 0, layoutOptions);

        // Save the workbook as an XLSX file
        workbook.Save(excelPath, SaveFormat.Xlsx);
    }
}