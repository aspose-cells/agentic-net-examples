// Title: Move an OleObject with Top/Left Pixel Offsets Using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, inserts an OLE object at row 5/column 5 (100 × 100 px), then shifts it down 30 px and right 50 px by adjusting its Top and Left properties before saving the file.
// Keywords: Aspose.Cells OleObject position | C# move OLE object top left | adjust OleObject coordinates | Aspose.Cells set OleObject Top | shift OleObject location Excel | Aspose.Cells example GitHub | pixel offset Excel shape
// Common Searches: how to change OleObject position Aspose.Cells C# | move OLE object down and right in Excel with Aspose | Aspose.Cells Top and Left properties example | C# code to offset OleObject location
// Developer Intent: Reposition a retrieved OleObject by adding pixel values to its Top and Left properties.
// Use Cases: Align an embedded chart with other worksheet shapes after insertion. | Apply consistent margins to a batch of OLE objects programmatically. | Adjust preview placement of linked documents based on user‑defined spacing.
// AI Prompts: Write C# code that moves an existing OleObject in a worksheet by given pixel offsets using Aspose.Cells. | Explain how the Top and Left properties of an OleObject map to its on‑sheet location and how to convert them to cell references. | Show how to keep OleObject position changes when saving a workbook in XLSX and XLS formats.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectPositionDemo
{
    // Creates a new workbook, inserts an OLE object at row 5/column 5 (100 × 100 px), then shifts it down 30 px and right 50 px by adjusting its Top and Left properties before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Prepare image data for the OLE object (placeholder image)
                byte[] imageData;
                const string imagePath = "sampleImage.png";

                if (File.Exists(imagePath))
                {
                    // Load existing image file
                    imageData = File.ReadAllBytes(imagePath);
                }
                else
                {
                    // Use a minimal 1x1 PNG (light gray) as placeholder
                    // Base64 for a 1x1 pixel PNG with RGB(211,211,211)
                    const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                    imageData = Convert.FromBase64String(base64Png);
                }

                // Add an OLE object at initial position (row 5, column 5) with size 100x100 pixels
                int oleIndex = worksheet.OleObjects.Add(5, 5, 100, 100, imageData);

                // Retrieve the added OLE object
                OleObject ole = worksheet.OleObjects[oleIndex];

                // Define offset values (in pixels)
                int topOffset = 30;   // move down by 30 pixels
                int leftOffset = 50;  // move right by 50 pixels

                // Adjust the position using the inherited Top and Left properties
                ole.Top += topOffset;
                ole.Left += leftOffset;

                // Save the workbook (lifecycle save)
                const string outputPath = "OleObjectMoved.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
