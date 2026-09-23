// Title: How to serialize a worksheet OleObject collection to JSON using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates over the first worksheet's OleObjectCollection, and generates a formatted JSON file containing each object's name, lock state, dimensions, position, and Base64‑encoded binary data. | Create a reusable method that maps an OleObject to a plain object, encodes its ObjectData property to Base64, and returns a JSON string using System.Text.Json. | Provide robust error‑handling for exporting embedded OLE objects to JSON, including checks for the source file's existence and safe writing of the output file.
// Common Searches: aspnet cells c# export embedded ole objects to json file | serialize oleobjectcollection to json using aspose.cells | c# get base64 data of excel OLE objects with Aspose | how to write ole object metadata from worksheet to json in .NET | asp.net core read ole objects from xlsx and output json
// Tags: Aspose.Cells OLE object JSON export | C# Base64 encoding of Excel embedded objects | worksheet OleObject property extraction | formatted JSON generation with System.Text.Json | robust file existence check in Aspose.Cells workflow

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject types

// The example loads an Excel workbook, accesses the OleObjectCollection on the first worksheet, gathers each object's name, lock status, size, position, and binary data (encoded as Base64), serializes the collection to an indented JSON string with System.Text.Json, and writes the result to 'OleObjects.json' while handling missing files and write errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "OleObjects.json";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of OLE objects on the worksheet
            OleObjectCollection oleObjects = worksheet.OleObjects;

            // Prepare a list of objects that can be serialized to JSON
            var oleList = new List<object>();

            foreach (OleObject ole in oleObjects)
            {
                // Gather available properties; omit those not present in older API versions
                var oleInfo = new
                {
                    ole.Name,
                    ole.IsLocked,
                    ole.Width,
                    ole.Height,
                    ole.Left,
                    ole.Top,
                    // Convert binary object data to Base64 string for JSON compatibility
                    ObjectData = ole.ObjectData != null ? Convert.ToBase64String(ole.ObjectData) : null
                };

                oleList.Add(oleInfo);
            }

            // Serialize the list to a formatted JSON string
            string json = JsonSerializer.Serialize(oleList, new JsonSerializerOptions { WriteIndented = true });

            // Write the JSON to a file for external auditing/integration
            try
            {
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"OLE object information has been written to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write JSON file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log unexpected errors without crashing the application
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
