using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace OleObjectsJsonExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Minimal 1x1 transparent PNG (hard‑coded) for the OLE object icon
                byte[] iconData = new byte[]
                {
                    0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                    0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                    0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                    0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                    0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                    0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
                    0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
                    0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                    0x42,0x60,0x82
                };

                // First OLE object (embedded) – using the dummy icon data
                worksheet.OleObjects.Add(2, 2, 150, 200, iconData);

                // Second OLE object (linked) – add only if the source file exists
                string linkedFile = "linkedSample.xlsx";
                if (File.Exists(linkedFile))
                {
                    worksheet.OleObjects.Add(5, 5, 120, 180, iconData, linkedFile);
                }

                // Retrieve the OleObject collection
                var oleObjects = worksheet.OleObjects;

                // Prepare a list of serializable objects containing relevant OLE information
                var oleInfoList = new List<object>();
                for (int i = 0; i < oleObjects.Count; i++)
                {
                    var ole = oleObjects[i];
                    oleInfoList.Add(new
                    {
                        Index = i,
                        UpperLeftRow = ole.UpperLeftRow,
                        UpperLeftColumn = ole.UpperLeftColumn,
                        Height = ole.Height,
                        Width = ole.Width,
                        FileFormatType = ole.FileFormatType.ToString(),
                        IsLink = ole.IsLink,
                        ObjectSourceFullName = ole.ObjectSourceFullName,
                        ProgID = ole.ProgID,
                        DisplayAsIcon = ole.DisplayAsIcon,
                        AutoLoad = ole.AutoLoad,
                        AutoUpdate = ole.AutoUpdate
                    });
                }

                // Serialize the list to JSON with indentation for readability
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(oleInfoList, jsonOptions);

                // Output JSON to console
                Console.WriteLine("Serialized OleObjectCollection:");
                Console.WriteLine(json);

                // Save JSON to a file for external auditing
                string jsonPath = "OleObjectsAudit.json";
                File.WriteAllText(jsonPath, json);
                Console.WriteLine($"JSON saved to: {Path.GetFullPath(jsonPath)}");

                // Save the workbook
                string workbookPath = "WorkbookWithOleObjects.xlsx";
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}