using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Loading;

class DbfJsonToCsvConverter
{
    static void Main()
    {
        // Path to the source JSON file that represents DBF data
        string jsonFilePath = "source.json";

        // Read the entire JSON content
        string jsonData = File.ReadAllText(jsonFilePath);

        // Create a new workbook (in-memory Excel file)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure JSON load options to preserve numeric and date types
        JsonLoadOptions jsonLoadOptions = new JsonLoadOptions
        {
            LayoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true,          // Treat JSON arrays as tables
                ConvertNumericOrDate = true,  // Convert numeric strings and dates to proper types
                DateFormat = "yyyy-MM-dd",    // Expected date format in JSON
                NumberFormat = "#,##0.00"     // Desired number format
            }
        };

        // Import JSON data into the worksheet starting at cell A1 (row 0, column 0)
        JsonUtility.ImportData(jsonData, worksheet.Cells, 0, 0, jsonLoadOptions.LayoutOptions);

        // Save the workbook as CSV, preserving data types (numeric and dates will be written in their native form)
        string csvOutputPath = "output.csv";
        workbook.Save(csvOutputPath, SaveFormat.Csv);

        Console.WriteLine($"Conversion completed. CSV file saved to: {csvOutputPath}");
    }
}