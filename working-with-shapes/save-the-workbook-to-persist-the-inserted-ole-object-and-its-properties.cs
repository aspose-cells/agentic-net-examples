using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Load icon image bytes if the file exists; otherwise use a 1x1 transparent PNG
            byte[] iconData;
            const string iconPath = "icon.png";
            if (File.Exists(iconPath))
            {
                iconData = File.ReadAllBytes(iconPath);
            }
            else
            {
                iconData = GetPlaceholderPng();
            }

            // Add an OLE object placeholder to the worksheet
            int oleIndex = sheet.OleObjects.Add(2, 2, 200, 200, iconData);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Load embedded Excel file bytes; if not found, create a temporary workbook in memory
            byte[] embeddedData;
            const string embedPath = "sample.xlsx";
            if (File.Exists(embedPath))
            {
                embeddedData = File.ReadAllBytes(embedPath);
            }
            else
            {
                Workbook tempWb = new Workbook();
                using (MemoryStream ms = new MemoryStream())
                {
                    tempWb.Save(ms, SaveFormat.Xlsx);
                    embeddedData = ms.ToArray();
                }
            }

            // Set OLE object data and properties
            ole.ObjectData = embeddedData;
            ole.ProgID = "Excel.Sheet";
            ole.DisplayAsIcon = true;
            ole.Label = "Embedded Sample.xlsx";
            ole.AutoLoad = true;

            // Save the workbook with the OLE object embedded
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                EmbedOoxmlAsOleObject = true
            };
            workbook.Save("OleObjectWorkbook.xlsx", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Returns a byte array representing a 1x1 transparent PNG
    private static byte[] GetPlaceholderPng()
    {
        // Base64 for a 1x1 transparent PNG
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XbZcAAAAASUVORK5CYII=";
        return Convert.FromBase64String(base64Png);
    }
}