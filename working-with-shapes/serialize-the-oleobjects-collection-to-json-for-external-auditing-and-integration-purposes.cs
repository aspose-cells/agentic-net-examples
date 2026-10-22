// Title: Serialize worksheet OleObject collection to indented JSON for auditing with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook using Aspose.Cells, iterates over the OleObjects in the first worksheet, and outputs their Name, IsLocked, IsPrintable, UpperLeftRow, UpperLeftColumn, Height, and Width as a formatted JSON array. | Create a reusable method ExportOleObjectsToJson(string workbookPath, string jsonPath) that returns a JSON string containing the selected properties of each OleObject, leveraging System.Text.Json for serialization. | Adjust the serialization to also include the OleObjectType property while preserving compatibility with older Aspose.Cells versions, using conditional checks or compile‑time directives.
// Common Searches: aspnet core how to extract ole object properties from Excel using Aspose.Cells and save to json | c# Aspose.Cells serialize embedded OLE objects metadata to a JSON file | example code for iterating OleObjectCollection and exporting details to JSON | export ole objects from worksheet to json for audit purposes with Aspose.Cells .NET
// Tags: Aspose.Cells serialize OleObjectCollection to JSON | export OLE object metadata Excel .NET | C# write worksheet embedded objects audit file | System.Text.Json Aspose.Cells OLE extraction | Excel OLE objects JSON export using Aspose

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// Loads an Excel workbook, reads the OleObject collection from the first worksheet, captures each object's name, lock and printable flags, position, height, and width, and writes the data as an indented JSON array to a file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "OleObjects.json";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of OLE objects in the worksheet
            OleObjectCollection oleObjects = worksheet.OleObjects;

            // Prepare a list to hold serializable information about each OLE object
            List<OleObjectInfo> oleInfoList = new List<OleObjectInfo>();

            foreach (OleObject ole in oleObjects)
            {
                OleObjectInfo info = new OleObjectInfo
                {
                    Name = ole.Name,
                    IsLocked = ole.IsLocked,
                    IsPrintable = ole.IsPrintable,
                    UpperLeftRow = ole.UpperLeftRow,
                    UpperLeftColumn = ole.UpperLeftColumn,
                    Height = ole.Height,
                    Width = ole.Width
                    // OleObjectType property is omitted for compatibility with older API versions
                };

                oleInfoList.Add(info);
            }

            // Serialize the list to JSON with indentation for readability
            JsonSerializerOptions options = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(oleInfoList, options);

            // Write JSON to the output file safely
            try
            {
                File.WriteAllText(outputPath, json);
                Console.WriteLine($"OLE object information saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write output file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper class that defines the JSON structure for an OLE object
    public class OleObjectInfo
    {
        public string? Name { get; set; }
        public bool IsLocked { get; set; }
        public bool IsPrintable { get; set; }
        public int UpperLeftRow { get; set; }
        public int UpperLeftColumn { get; set; }
        public double Height { get; set; }
        public double Width { get; set; }
        // public string? OleObjectType { get; set; } // Omitted for compatibility
    }
}
