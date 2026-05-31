using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class JsonToCsvConverter
{
    static void Main()
    {
        try
        {
            // Path to the source JSON file
            string jsonFilePath = "data.json";

            // Verify that the JSON file exists
            if (!File.Exists(jsonFilePath))
                throw new FileNotFoundException($"JSON file not found: {jsonFilePath}");

            // Desired delimiter for the CSV output (e.g., semicolon)
            string customDelimiter = ";";

            // Read the entire JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Configure JSON layout options to treat arrays as tables
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, cells, 0, 0, layoutOptions);

            // Prepare CSV save options with the custom delimiter (Separator expects a char)
            TxtSaveOptions csvSaveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = customDelimiter[0]   // use first character of the delimiter string
            };

            // Save the workbook as a CSV file using the specified options (save rule)
            workbook.Save("output.csv", csvSaveOptions);
        }
        catch (FileNotFoundException fnfEx)
        {
            Console.WriteLine($"File error: {fnfEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}