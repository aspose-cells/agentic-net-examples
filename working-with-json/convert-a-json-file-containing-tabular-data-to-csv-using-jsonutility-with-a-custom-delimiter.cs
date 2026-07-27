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
            // Paths for input JSON and output CSV files
            string jsonFilePath = "input.json";
            string csvFilePath = "output.csv";

            // Verify that the JSON input file exists
            if (!File.Exists(jsonFilePath))
                throw new FileNotFoundException($"Input JSON file not found: {jsonFilePath}");

            // Custom delimiter for the CSV output (must be a char)
            char customDelimiter = '|';

            // Read the entire JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new workbook (Excel file in memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set JSON layout options to treat arrays as tables (tabular data)
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import the JSON data into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, cells, 0, 0, layoutOptions);

            // Configure CSV save options with the custom delimiter
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = customDelimiter
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(csvFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook as a CSV file using the specified options
            workbook.Save(csvFilePath, csvOptions);

            Console.WriteLine($"JSON data has been converted to CSV at '{csvFilePath}' with delimiter '{customDelimiter}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}