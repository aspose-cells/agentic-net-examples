// Title: Assign a custom label to an OLE object in Excel using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert an OLE object with a 1×1 transparent PNG icon, set the OleObject.Label property, save the file, reload it, and verify that the label persists.
// Keywords: Aspose.Cells | C# OleObject label | Excel OLE object custom label | OleObject.Label property | add OLE object Aspose.Cells | placeholder PNG for OLE | read OLE label after save
// Common Searches: how to set OleObject.Label in Aspose.Cells C# | change display label of Excel OLE object programmatically | Aspose.Cells add OLE object with custom icon | verify OLE object label after saving workbook | C# example for labeling OLE objects in Excel
// Developer Intent: The developer needs to assign or modify the display label of an OLE object embedded in an Excel workbook using the Aspose.Cells .NET API.
// Use Cases: Insert a new OLE object with a custom label and placeholder image when generating reports. | Open an existing spreadsheet, locate OLE objects, update their labels to reflect current content, and save the changes. | Iterate through all OLE objects on a worksheet to assign unique, descriptive labels for downstream processing.
// AI Prompts: Generate C# code that adds an OLE object with a custom label and a transparent PNG using Aspose.Cells. | Show how to load an existing Excel file, change the Label of each OleObject, and persist the modifications. | Explain how to confirm that an OLE object's label was saved correctly and can be read back after workbook serialization.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectDemo
{
    // Demonstrates how to create a workbook, insert an OLE object with a 1×1 transparent PNG icon, set the OleObject.Label property, save the file, reload it, and verify that the label persists.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Generate a simple placeholder image for the OLE object's icon
                byte[] imageData = CreatePlaceholderImage();

                // Add an OLE object with the placeholder image
                int oleIndex = worksheet.OleObjects.Add(10, 10, 200, 200, imageData);
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Assign a custom display label to the OLE object
                oleObject.Label = "My Custom OLE Label";

                // Save the workbook to a file
                string outputPath = "OleObjectLabelDemo.xlsx";
                workbook.Save(outputPath);

                // Verify the label was saved
                if (File.Exists(outputPath))
                {
                    Workbook loadedWorkbook = new Workbook(outputPath);
                    OleObject loadedOleObject = loadedWorkbook.Worksheets[0].OleObjects[0];
                    Console.WriteLine("OLE Object Label: " + loadedOleObject.Label);
                }
                else
                {
                    Console.WriteLine("Failed to save the workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        // Returns a 1x1 transparent PNG image as a byte array
        private static byte[] CreatePlaceholderImage()
        {
            // Base64 representation of a 1x1 transparent PNG
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            return Convert.FromBase64String(base64Png);
        }
    }
}
