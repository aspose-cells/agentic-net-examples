// Title: C# – Assign a custom display label to an OLE object with Aspose.Cells
// Description: Demonstrates how to create a workbook, embed an OLE object with a placeholder PNG icon, set the OleObject.Label property to a custom string, save the file, reload it, and output the stored label to confirm persistence.
// Keywords: Aspose.Cells OLE label C# | OleObject.Label property | add OLE object Aspose.Cells | custom OLE caption Excel | verify OLE label after save | placeholder PNG icon OLE | C# Excel OLE embedding
// Common Searches: how to set OLE object label in Aspose.Cells .NET | C# example for OleObject.Label | change display caption of embedded OLE in Excel using Aspose | retrieve OLE object label after workbook save | add OLE object with custom icon and label programmatically
// Developer Intent: Add an OLE object to a worksheet and define a user‑friendly label that persists after saving.
// Use Cases: Generate reports that embed documents or charts with clear captions for end users. | Create templates where each OLE object needs a descriptive label displayed in Excel. | Load existing workbooks and read OLE object labels to populate a custom dashboard.
// AI Prompts: Show C# code to embed an OLE object with a custom label and icon using Aspose.Cells. | Provide a snippet that updates the Label of an existing OLE object in a saved workbook and handles errors. | Explain how to enumerate all OLE objects in a worksheet and print their labels with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectLabelDemo
{
    // Demonstrates how to create a workbook, embed an OLE object with a placeholder PNG icon, set the OleObject.Label property to a custom string, save the file, reload it, and output the stored label to confirm persistence.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare a simple placeholder PNG image (1x1 transparent pixel) for the OLE object's icon.
                byte[] imageData = GetPlaceholderPng();

                // Add an OLE object to the worksheet.
                // Parameters: upper row, left column, height (pixels), width (pixels), image data for the icon.
                int oleIndex = worksheet.OleObjects.Add(10, 10, 200, 200, imageData);
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Assign a custom display label.
                oleObject.Label = "Custom OLE Label";

                // Save the workbook.
                string filePath = "OleObjectLabelDemo.xlsx";
                workbook.Save(filePath);

                // Verify the label by loading the workbook back.
                if (File.Exists(filePath))
                {
                    try
                    {
                        Workbook loadedWorkbook = new Workbook(filePath);
                        OleObject loadedOle = loadedWorkbook.Worksheets[0].OleObjects[0];
                        Console.WriteLine("OLE Object Label: " + loadedOle.Label);
                    }
                    catch (Exception loadEx)
                    {
                        Console.WriteLine("Error loading workbook: " + loadEx.Message);
                    }
                }
                else
                {
                    Console.WriteLine("Error: Saved file not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Returns a byte array containing a minimal PNG image (1x1 transparent pixel).
        private static byte[] GetPlaceholderPng()
        {
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            return Convert.FromBase64String(base64Png);
        }
    }
}
