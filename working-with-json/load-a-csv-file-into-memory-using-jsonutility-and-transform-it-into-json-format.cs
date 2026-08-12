// Title: C# – Convert CSV to JSON using Aspose.Cells JsonUtility
// Description: Loads a CSV file into an Aspose.Cells workbook, creates a range covering all data, configures JsonSaveOptions (header row, empty cells), exports the range to a JSON string with JsonUtility.ExportRangeToJson, and saves the result to a file.
// Keywords: Aspose.Cells CSV to JSON | JsonUtility ExportRangeToJson C# | ImportCSV Aspose.Cells example | JsonSaveOptions header row | C# convert CSV file to JSON | Aspose.Cells write JSON file
// Common Searches: Aspose.Cells convert CSV to JSON C# | JsonUtility ExportRangeToJson usage | ImportCSV then export JSON Aspose.Cells | C# code to export worksheet range as JSON | Save JSON output from Aspose.Cells
// Developer Intent: Read a CSV file, load it into an Aspose.Cells workbook, and export the worksheet data as a JSON string or file.
// Use Cases: Transform flat CSV data into JSON for web APIs. | Generate configuration files from spreadsheet data without manual conversion. | Create JSON payloads for services by exporting selected worksheet ranges.
// AI Prompts: Show a C# snippet that reads a CSV with Aspose.Cells, sets JsonSaveOptions, and writes the exported JSON to a file. | Explain how to adjust JsonSaveOptions to produce nested JSON structures when exporting a range. | Provide guidance on streaming large CSV files into Aspose.Cells before converting them to JSON for optimal performance.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads a CSV file into an Aspose.Cells workbook, creates a range covering all data, configures JsonSaveOptions (header row, empty cells), exports the range to a JSON string with JsonUtility.ExportRangeToJson, and saves the result to a file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.Error.WriteLine($"Error: CSV file not found at path '{csvPath}'.");
                return;
            }

            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Get the Cells collection of the first worksheet
            Cells cells = workbook.Worksheets[0].Cells;

            // Import CSV data into the worksheet starting at cell A1
            // Using comma as delimiter and converting numeric strings to numbers
            cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Determine the used range dimensions
            int lastRow = cells.MaxDataRow;          // zero‑based index of the last row with data
            int lastColumn = cells.MaxDataColumn;    // zero‑based index of the last column with data

            // Create a range that covers all imported data
            Aspose.Cells.Range dataRange = cells.CreateRange(0, 0, lastRow + 1, lastColumn + 1);

            // Configure JSON export options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true,
                ExportNestedStructure = false
            };

            // Export the range to a JSON string using the JsonUtility method
            string jsonResult = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);

            // Output the JSON string to the console
            Console.WriteLine(jsonResult);

            // Optionally, write the JSON string to a file
            string outputPath = "output.json";
            File.WriteAllText(outputPath, jsonResult);
            Console.WriteLine($"JSON output written to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
