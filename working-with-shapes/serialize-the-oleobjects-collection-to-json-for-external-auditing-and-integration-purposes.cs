// Title: C# – Export Excel OleObjects Collection to JSON using Aspose.Cells
// Description: This C# example creates a workbook, adds embedded and linked OLE objects, extracts their properties (position, size, format, link status, source path, ProgID, and Base64‑encoded data) from the worksheet's OleObjectCollection, and writes the information to a formatted JSON file with System.Text.Json. Ideal for auditing, integration, or migration scenarios.
// Keywords: Aspose.Cells | C# | .NET | OleObject | OleObjectCollection | JSON export | serialize OLE objects | Base64 | embedded OLE | linked OLE | Excel automation | audit Excel objects
// Common Searches: export oleobjects to json aspose.cells | serialize worksheet oleobject collection c# | extract ole object properties aspose.cells .net | save ole object metadata as json | convert ole object data to base64 c#
// Developer Intent: Generate a JSON file containing all OLE objects from a worksheet for external audit or integration.
// Use Cases: Maintain an audit log of embedded and linked OLE objects across workbooks | Feed OLE object metadata to a document‑management API | Validate linked OLE sources before publishing a spreadsheet | Migrate OLE objects to another system by reconstructing them from JSON
// AI Prompts: Write C# code that reads the generated OleObjects.json and recreates the OLE objects in a new workbook using Aspose.Cells. | Provide a LINQ query to filter only linked OLE objects from the JSON and save them to a separate file. | Create a unit test that confirms the Base64ObjectData field decodes to the original binary of an embedded OLE object. | Suggest how to stream large OLE object data to JSON without loading the entire workbook into memory.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject and OleObjectCollection

namespace OleObjectsJsonExport
{
    // DTO for JSON serialization; string properties are nullable to satisfy compiler warnings
    // This C# example creates a workbook, adds embedded and linked OLE objects, extracts their properties (position, size, format, link status, source path, ProgID, and Base64‑encoded data) from the worksheet's OleObjectCollection, and writes the information to a formatted JSON file with System.Text.Json. Ideal for auditing, integration, or migration scenarios.
    public class OleObjectInfo
    {
        public int UpperLeftRow { get; set; }
        public int UpperLeftColumn { get; set; }
        public int Height { get; set; }
        public int Width { get; set; }
        public string? FileFormatType { get; set; }
        public bool IsLink { get; set; }
        public string? ObjectSourceFullName { get; set; }
        public string? ProgID { get; set; }
        public string? Base64ObjectData { get; set; }   // Embedded binary data as Base64 string
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Prepare a placeholder icon (empty byte array) for OLE objects
                // -----------------------------------------------------------------
                byte[] iconBytes = Array.Empty<byte>();

                // -----------------------------------------------------------------
                // 3. Add sample OLE objects (embedded and linked)
                // -----------------------------------------------------------------
                // Embedded OLE object (using the placeholder icon; no source file needed for demo)
                int embeddedIndex = sheet.OleObjects.Add(2, 2, 100, 100, iconBytes);
                OleObject ole1 = sheet.OleObjects[embeddedIndex];

                // Linked OLE object – ensure the linked file exists
                string linkedFilePath = Path.Combine(Path.GetTempPath(), "linkedFile.docx");
                if (!File.Exists(linkedFilePath))
                {
                    File.WriteAllText(linkedFilePath, "Dummy linked file content");
                }
                int linkedIndex = sheet.OleObjects.Add(5, 5, 120, 120, iconBytes, linkedFilePath);
                OleObject ole2 = sheet.OleObjects[linkedIndex];

                // -----------------------------------------------------------------
                // 4. Extract information from the OleObjectCollection
                // -----------------------------------------------------------------
                OleObjectCollection oleCollection = sheet.OleObjects;
                List<OleObjectInfo> oleInfoList = new List<OleObjectInfo>();

                for (int i = 0; i < oleCollection.Count; i++)
                {
                    OleObject ole = oleCollection[i];
                    OleObjectInfo info = new OleObjectInfo
                    {
                        UpperLeftRow = ole.UpperLeftRow,
                        UpperLeftColumn = ole.UpperLeftColumn,
                        Height = ole.Height,
                        Width = ole.Width,
                        FileFormatType = ole.FileFormatType.ToString(),
                        IsLink = ole.IsLink,
                        ObjectSourceFullName = ole.ObjectSourceFullName,
                        ProgID = ole.ProgID,
                        Base64ObjectData = ole.ObjectData != null ? Convert.ToBase64String(ole.ObjectData) : null
                    };
                    oleInfoList.Add(info);
                }

                // -----------------------------------------------------------------
                // 5. Serialize the list to JSON using System.Text.Json
                // -----------------------------------------------------------------
                JsonSerializerOptions jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true // pretty‑print for readability
                };
                string json = JsonSerializer.Serialize(oleInfoList, jsonOptions);

                // -----------------------------------------------------------------
                // 6. Save the JSON to a file
                // -----------------------------------------------------------------
                string outputPath = "OleObjects.json";
                File.WriteAllText(outputPath, json);

                Console.WriteLine($"OLE objects information exported to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
