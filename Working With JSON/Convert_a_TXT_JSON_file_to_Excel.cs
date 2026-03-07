using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class TxtJsonToExcel
{
    static void Main()
    {
        // Path to the source TXT file that contains JSON data
        string txtPath = "input.txt";

        // Path where the resulting Excel file will be saved
        string excelPath = "output.xlsx";

        // Read the entire JSON string from the text file
        string jsonContent = File.ReadAllText(txtPath);

        // Create a new workbook (empty Excel file)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure import options (optional)
        JsonLayoutOptions layoutOptions = new JsonLayoutOptions
        {
            // Treat JSON arrays as tables so each element becomes a row
            ArrayAsTable = true
        };

        // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonContent, cells, 0, 0, layoutOptions);

        // Save the workbook as an XLSX file
        workbook.Save(excelPath, SaveFormat.Xlsx);

        Console.WriteLine($"Conversion completed successfully. Excel file saved to: {excelPath}");
    }
}