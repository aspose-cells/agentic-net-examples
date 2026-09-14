// Title: Export a named range from an Excel workbook to a JSON file with cell addresses and values using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, retrieves a specific named range, and creates a JSON array where each element contains the cell's A1 address and its value. | Generate a .NET console program that validates the input workbook, extracts all cells from a given named range, and serializes the address/value pairs to a formatted JSON file. | Provide a C# method that iterates over an Aspose.Cells named range and returns a JSON string representing each cell's address and value.
// Common Searches: how to read a named range with Aspose.Cells and export it to JSON in C# | Aspose.Cells C# convert named range cells to JSON array of address/value | C# export Excel named range data to JSON file using Aspose.Cells | serialize Aspose.Cells named range to JSON with cell addresses
// Tags: Aspose.Cells named range JSON export | C# serialize Excel range to JSON | extract cell address value pairs .NET | Aspose.Cells range iteration for JSON output | convert Excel named range to JSON array

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example checks for the existence of an input workbook, loads it with Aspose.Cells, retrieves a named range called "MyRange", iterates through each cell to collect its A1 address and value, serializes the collection into an indented JSON array, and writes the result to "namedRange.json" while handling loading and I/O errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "namedRange.json";
            const string rangeName = "MyRange";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Error loading workbook: {loadEx.Message}");
                return;
            }

            // Retrieve the named range; GetRangeByName returns an Aspose.Cells.Range object
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName(rangeName);
            if (namedRange == null)
            {
                Console.WriteLine($"Error: Named range \"{rangeName}\" does not exist in the workbook.");
                return;
            }

            // Get the worksheet that contains the range
            Worksheet ws = namedRange.Worksheet;
            if (ws == null)
            {
                Console.WriteLine("Error: Unable to determine the worksheet for the named range.");
                return;
            }

            // Prepare a list to hold address/value pairs for JSON serialization
            List<Dictionary<string, object>> jsonItems = new List<Dictionary<string, object>>();

            // Iterate through each cell in the range
            for (int i = 0; i < namedRange.RowCount; i++)
            {
                for (int j = 0; j < namedRange.ColumnCount; j++)
                {
                    Cell cell = ws.Cells[namedRange.FirstRow + i, namedRange.FirstColumn + j];
                    var item = new Dictionary<string, object>
                    {
                        ["address"] = cell.Name,   // A1 style address
                        ["value"] = cell.Value     // Cell value (object)
                    };
                    jsonItems.Add(item);
                }
            }

            // Serialize the list to a formatted JSON string
            string json = JsonSerializer.Serialize(jsonItems, new JsonSerializerOptions { WriteIndented = true });

            // Ensure the output directory exists and write the JSON file
            try
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                File.WriteAllText(outputPath, json);
                Console.WriteLine($"Named range data has been written to \"{outputPath}\".");
            }
            catch (Exception ioEx)
            {
                Console.WriteLine($"Error writing JSON file: {ioEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
