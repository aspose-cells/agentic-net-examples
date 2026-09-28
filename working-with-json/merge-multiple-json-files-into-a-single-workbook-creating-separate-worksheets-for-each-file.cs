// Title: Combine multiple JSON files into a single Excel workbook with separate worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that scans a directory for *.json files, creates a new worksheet in an Aspose.Cells workbook for each file, and names the worksheet after the source filename. | Write a method that receives a JSON array string and populates an Aspose.Cells worksheet with column headers derived from the first object and data rows for each element. | Provide C# logic to save the populated Aspose.Cells workbook as an XLSX file to a given output path, creating the output folder if it does not exist.
// Common Searches: c# aspocells read all json files in a folder and export each to a separate worksheet | how to create a new worksheet for each json file using Aspose.Cells .NET | import json array into Excel with Aspose.Cells and auto‑fit columns | merge multiple json datasets into one Excel workbook with Aspose.Cells C# example | aspocells convert json files to xlsx with worksheet names from filenames
// Tags: aspocells import json array to worksheet | aspocells create workbook from multiple json files | aspocells auto‑fit columns after json import | c# convert json files to xlsx using Aspose.Cells | aspocells generate worksheet names from filenames

using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

// The example scans a specified folder for *.json files, creates a new Aspose.Cells Workbook, adds a worksheet for each file named after the file (without extension), parses each JSON array to write column headers and rows, auto‑fits columns, and finally saves the combined workbook as an XLSX file.
class JsonToWorkbookMerger
{
    static void Main()
    {
        try
        {
            // Folder containing the JSON files
            string jsonFolderPath = @"C:\Data\JsonFiles";

            // Verify the JSON folder exists
            if (!Directory.Exists(jsonFolderPath))
            {
                Console.WriteLine($"JSON folder not found: {jsonFolderPath}");
                return;
            }

            // Output Excel file path
            string outputExcelPath = @"C:\Data\MergedWorkbook.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputExcelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Get all JSON files in the specified folder
            string[] jsonFiles = Directory.GetFiles(jsonFolderPath, "*.json");

            foreach (string jsonFilePath in jsonFiles)
            {
                try
                {
                    // Read the JSON content from the file
                    string jsonContent = File.ReadAllText(jsonFilePath);

                    // Add a new worksheet for this JSON file
                    Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
                    sheet.Name = Path.GetFileNameWithoutExtension(jsonFilePath);

                    // Import JSON data into the worksheet
                    ImportJsonToWorksheet(jsonContent, sheet);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process '{jsonFilePath}': {ex.Message}");
                }
            }

            // Save the merged workbook to the specified path
            workbook.Save(outputExcelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputExcelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Parses JSON (expected to be an array of objects) and writes it to the worksheet.
    private static void ImportJsonToWorksheet(string jsonContent, Worksheet sheet)
    {
        using JsonDocument doc = JsonDocument.Parse(jsonContent);
        JsonElement root = doc.RootElement;

        if (root.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("Root JSON element must be an array of objects.");
        }

        // Collect headers from the first object
        List<string> headers = new List<string>();
        foreach (JsonProperty prop in root[0].EnumerateObject())
        {
            headers.Add(prop.Name);
        }

        // Write headers (first row)
        for (int col = 0; col < headers.Count; col++)
        {
            sheet.Cells[0, col].PutValue(headers[col]);
        }

        // Write data rows
        int rowIndex = 1;
        foreach (JsonElement element in root.EnumerateArray())
        {
            for (int col = 0; col < headers.Count; col++)
            {
                string header = headers[col];
                if (element.TryGetProperty(header, out JsonElement valueElement))
                {
                    object value = GetJsonValue(valueElement);
                    sheet.Cells[rowIndex, col].PutValue(value);
                }
                else
                {
                    sheet.Cells[rowIndex, col].PutValue(string.Empty);
                }
            }
            rowIndex++;
        }

        // Auto-fit columns for better appearance
        sheet.AutoFitColumns();
    }

    // Converts a JsonElement to a .NET primitive suitable for Aspose.Cells.
    private static object GetJsonValue(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number => element.TryGetInt64(out long l) ? (object)l :
                                     element.TryGetDouble(out double d) ? d : element.GetRawText(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null,
            _ => element.GetRawText()
        };
    }
}
