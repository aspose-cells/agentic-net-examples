// Title: C# – Export Aspose.Cells Worksheet OleObjects to JSON for Auditing & Integration
// Description: Shows how to create a workbook, add embedded and linked OLE objects, read each object's key properties (row, column, size, format, link status, source file, ProgID, display‑as‑icon) and serialize the OleObjects collection to a formatted JSON file for external audit, reporting, or downstream processing.
// Keywords: Aspose.Cells | OleObject | C# JSON serialization | export OLE objects | Excel audit | worksheet OleObjects collection | placeholder image | embedded OLE | linked OLE | metadata extraction | Aspose.Cells API | JSON file output
// Common Searches: How to export OleObjects from Aspose.Cells to JSON | C# serialize worksheet OLE objects to JSON | Aspose.Cells get OLE object properties | Export embedded and linked OLE objects for audit | Save OLE object metadata as JSON file
// Developer Intent: Create a JSON file that lists all OleObject items in a worksheet with their essential attributes for auditing or integration with other systems.
// Use Cases: Produce an audit log of every embedded or linked OLE object in generated Excel files. | Feed OLE object metadata to a downstream service that consumes JSON for validation or reporting. | Verify placement, dimensions, and link status of OLE objects before publishing a workbook.
// AI Prompts: Write C# code that reads the generated OleObjects.json and reconstructs OleObjectInfo objects for further analysis. | Provide a method that filters the exported OLE objects by IsLink and writes the filtered set to a new JSON file. | Show how to add comprehensive error handling for missing placeholder images when adding OLE objects with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectsJsonExport
{
    // Simple DTO to hold relevant OLE object information for JSON serialization
    // Shows how to create a workbook, add embedded and linked OLE objects, read each object's key properties (row, column, size, format, link status, source file, ProgID, display‑as‑icon) and serialize the OleObjects collection to a formatted JSON file for external audit, reporting, or downstream processing.
    public class OleObjectInfo
    {
        public int UpperLeftRow { get; set; }
        public int UpperLeftColumn { get; set; }
        public int Height { get; set; }
        public int Width { get; set; }
        public string FileFormatType { get; set; }
        public bool IsLink { get; set; }
        public string ObjectSourceFullName { get; set; }
        public string ProgID { get; set; }
        public bool DisplayAsIcon { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Placeholder image (1x1 pixel PNG) required by Aspose.Cells when adding OLE objects
                byte[] placeholderImage = GetPlaceholderImageBytes();

                // Add sample OLE objects to the worksheet
                // Embedded OLE object (no source file)
                worksheet.OleObjects.Add(2, 2, 150, 200, placeholderImage);

                // Linked OLE object (source file name provided)
                worksheet.OleObjects.Add(5, 5, 120, 180, placeholderImage, "sample.docx");

                // Prepare a list to hold serializable information about each OLE object
                List<OleObjectInfo> oleInfoList = new List<OleObjectInfo>();

                // Iterate through the OleObjectCollection
                foreach (OleObject ole in worksheet.OleObjects)
                {
                    oleInfoList.Add(new OleObjectInfo
                    {
                        UpperLeftRow = ole.UpperLeftRow,
                        UpperLeftColumn = ole.UpperLeftColumn,
                        Height = ole.Height,
                        Width = ole.Width,
                        FileFormatType = ole.FileFormatType.ToString(),
                        IsLink = ole.IsLink,
                        ObjectSourceFullName = ole.ObjectSourceFullName,
                        ProgID = ole.ProgID,
                        DisplayAsIcon = ole.DisplayAsIcon
                    });
                }

                // Serialize the list to JSON with indentation for readability
                string json = JsonSerializer.Serialize(oleInfoList, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // Save the JSON to a file (external auditing/integration)
                string outputPath = "OleObjects.json";
                File.WriteAllText(outputPath, json);

                Console.WriteLine($"OLE objects information exported to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Returns a valid PNG image byte array (1x1 pixel) used as a placeholder for OLE object icons
        private static byte[] GetPlaceholderImageBytes()
        {
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
            return Convert.FromBase64String(base64Png);
        }
    }
}
