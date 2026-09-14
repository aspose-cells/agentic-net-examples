// Title: Load a CSV file into an Aspose.Cells Workbook and export it as a JSON string using JsonSaveOptions in C#
// AI Prompts: Generate C# code that reads a CSV file into an Aspose.Cells Workbook, then uses JsonSaveOptions to write the workbook to a MemoryStream and returns the JSON string. | Modify the program to accept the input CSV path and an optional output JSON file path from command‑line arguments, writing the JSON result to the specified file. | Add error handling that validates the CSV file exists, catches malformed CSV content, and ensures the JSON output is encoded in UTF‑8.
// Common Searches: c# how to read a CSV file with Aspose.Cells and convert it to JSON | using Aspose.Cells JsonSaveOptions to export workbook to JSON string | convert CSV to JSON in memory stream Aspose.Cells C# example | Aspose.Cells load CSV and save as JSON without creating intermediate Excel file
// Tags: Aspose.Cells load CSV workbook | JsonSaveOptions export to JSON | C# convert CSV to JSON in memory | Aspose.Cells memory stream JSON output | CSV to JSON conversion using Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Demonstrates loading a CSV file into an Aspose.Cells Workbook, exporting the workbook to a MemoryStream with JsonSaveOptions, converting the stream to a UTF‑8 JSON string, and printing the result, with basic file‑existence checking and exception handling.
class CsvToJsonConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.csv";

            // Ensure the input CSV file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the CSV file into a Workbook object
            Workbook workbook = new Workbook(inputPath);

            // Export the workbook to JSON using JsonSaveOptions
            string json;
            using (MemoryStream ms = new MemoryStream())
            {
                JsonSaveOptions jsonOptions = new JsonSaveOptions();
                workbook.Save(ms, jsonOptions);
                json = Encoding.UTF8.GetString(ms.ToArray());
            }

            // Output the resulting JSON
            Console.WriteLine(json);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
