// Title: Persist Embedded OLE Object and Its Properties When Saving an Aspose.Cells Workbook (C#)
// Description: Demonstrates how to add an OLE object with a custom icon to a worksheet, configure ProgID, label, DisplayAsIcon and AutoLoad, and then save the workbook so the OLE object and all settings are retained. Includes a second save using OoxmlSaveOptions with EmbedOoxmlAsOleObject to store the entire OOXML package as an OLE object.
// Keywords: Aspose.Cells | C# | OLE object | embed Excel file | save workbook | OoxmlSaveOptions | DisplayAsIcon | AutoLoad | ProgID | embedded OOXML
// Common Searches: Aspose.Cells save workbook with embedded OLE object | C# add OLE object to worksheet Aspose.Cells | persist OLE object properties after saving | OoxmlSaveOptions EmbedOoxmlAsOleObject example | display OLE as icon Aspose.Cells
// Developer Intent: The developer needs to save a workbook while ensuring that the inserted OLE object and all its configured properties remain intact.
// Use Cases: Generate a report that contains a template Excel file as an OLE object, preserving the icon and auto‑load behavior after export. | Create a spreadsheet that shows a custom document icon, loads the linked file on open, and retains these settings when the file is shared. | Store the full OOXML package inside an OLE object for later extraction, using OoxmlSaveOptions during the save operation.
// AI Prompts: Write C# code with Aspose.Cells to insert an OLE object, set a custom icon, ProgID, label, and AutoLoad, then save the workbook preserving these settings. | Show how to use OoxmlSaveOptions in Aspose.Cells to embed the workbook's OOXML as an OLE object while saving. | Provide robust error handling for missing icon or embedded files when adding an OLE object with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add an OLE object with a custom icon to a worksheet, configure ProgID, label, DisplayAsIcon and AutoLoad, and then save the workbook so the OLE object and all settings are retained. Includes a second save using OoxmlSaveOptions with EmbedOoxmlAsOleObject to store the entire OOXML package as an OLE object.
class SaveOleObjectDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Paths for the icon image and the file to embed as OLE object
            string iconPath = "icon.png";
            string embedFilePath = "sample.xlsx";

            // Ensure the icon file exists – create a minimal 1x1 PNG if missing
            if (!File.Exists(iconPath))
            {
                // Transparent 1x1 PNG (base64 encoded)
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(base64Png);
                File.WriteAllBytes(iconPath, pngBytes);
            }

            // Ensure the file to embed exists – create an empty workbook if missing
            if (!File.Exists(embedFilePath))
            {
                new Workbook().Save(embedFilePath);
            }

            // Read the icon image data
            byte[] imageData = File.ReadAllBytes(iconPath);

            // Add an OLE object to the worksheet (initially with the icon image)
            int oleIndex = sheet.OleObjects.Add(2, 2, 200, 200, imageData);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Set the embedded file data and other properties
            ole.ObjectData = File.ReadAllBytes(embedFilePath);
            ole.ProgID = "Excel.Sheet";
            ole.DisplayAsIcon = true;
            ole.Label = "Embedded Sample";
            ole.AutoLoad = true; // Load automatically when the workbook is opened

            // Save the workbook using the standard Save method
            workbook.Save("OleObjectResult.xlsx");

            // Optionally, save with OoxmlSaveOptions to embed OOXML as an OLE object
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                EmbedOoxmlAsOleObject = true
            };
            workbook.Save("OleObjectResult_OoxmlEmbedded.xlsx", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
