// Title: Convert a JSON file to a CSV file with a semicolon delimiter using Aspose.Cells JsonUtility in C#
// AI Prompts: Generate C# code that reads a JSON file, imports the data into an Aspose.Cells workbook with JsonUtility, and exports it to a CSV file using a semicolon as the field separator. | Show how to configure TxtSaveOptions to specify a custom delimiter when saving a workbook as CSV with Aspose.Cells. | Add robust error handling that checks for the JSON input file's existence and captures exceptions during the JSON‑to‑CSV conversion.
// Common Searches: c# Aspose.Cells convert json to csv with custom delimiter | how to set semicolon separator when exporting workbook to csv using Aspose.Cells | JsonUtility import json data into worksheet then save as csv Aspose.Cells example | error handling for missing json file Aspose.Cells JsonUtility conversion | save workbook as csv with custom separator using TxtSaveOptions in .NET
// Tags: Aspose.Cells JsonUtility import JSON to worksheet | Aspose.Cells TxtSaveOptions custom CSV delimiter | JSON to CSV conversion C# Aspose.Cells | file existence validation Aspose.Cells conversion | exception handling Aspose.Cells JSON import

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// The sample reads a JSON file, uses Aspose.Cells JsonUtility to load the tabular data into a workbook, and saves the workbook as a CSV file with a semicolon delimiter via TxtSaveOptions, while validating the input file and handling potential errors.
class JsonToCsvConverter
{
    static void Main()
    {
        // Paths to the input JSON file and the output CSV file
        string jsonFilePath = "data.json";
        string csvFilePath = "output.csv";

        // Define the custom delimiter for the CSV output
        char customDelimiter = ';';

        try
        {
            // Verify that the JSON input file exists
            if (!File.Exists(jsonFilePath))
                throw new FileNotFoundException($"Input JSON file not found: {jsonFilePath}");

            // Read the entire JSON content from the file
            string jsonContent = File.ReadAllText(jsonFilePath);

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Prepare JSON layout options (first row contains column names)
            JsonLayoutOptions jsonOptions = new JsonLayoutOptions();
            // jsonOptions.FirstRowIsColumnNames = true; // Uncomment if supported by your version

            // Import the JSON data into the first worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, workbook.Worksheets[0].Cells, 0, 0, jsonOptions);

            // Set up CSV save options with the custom delimiter using TxtSaveOptions
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
            {
                Separator = customDelimiter
            };

            // Save the workbook as a CSV file using the specified options
            workbook.Save(csvFilePath, csvOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
