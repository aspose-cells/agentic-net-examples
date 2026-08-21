// Title: Lock OleObject Aspect Ratio in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds an OleObject with a PNG preview, sets IsAspectRatioLocked to true so the object retains its original width‑height ratio during resizing, and saves the file.
// Keywords: Aspose.Cells | OleObject | IsAspectRatioLocked | aspect ratio lock | C# | preserve proportions | resize without distortion | image preview | Excel automation
// Common Searches: Aspose.Cells lock OleObject aspect ratio | Set IsAspectRatioLocked C# | Keep OleObject proportions when resizing | Add OleObject with preview image Aspose.Cells | C# example for locking OleObject size
// Developer Intent: Enable the IsAspectRatioLocked property on an OleObject so its original width‑height ratio remains unchanged during any resize operation.
// Use Cases: Insert a company logo as an OleObject and ensure it scales proportionally across different worksheet layouts. | Embed a chart preview image as an OleObject while preserving its aspect ratio when users adjust rows or columns. | Generate automated reports where embedded documents must retain their original dimensions when the workbook is printed or exported.
// AI Prompts: Generate C# code using Aspose.Cells to add an OleObject with a PNG preview and lock its aspect ratio. | Explain how the IsAspectRatioLocked property affects OleObject resizing in Aspose.Cells and show a short example. | Provide step‑by‑step instructions to verify that an OleObject's aspect ratio is locked after insertion.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds an OleObject with a PNG preview, sets IsAspectRatioLocked to true so the object retains its original width‑height ratio during resizing, and saves the file.
class SetOleObjectAspectRatio
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image that will be used as the OleObject's preview
            string imagePath = "OleObjectPreview.png";

            if (File.Exists(imagePath))
            {
                // Load image data
                byte[] imageData = File.ReadAllBytes(imagePath);

                // Add an OleObject to the worksheet (row 5, column 2, height 200px, width 300px)
                int oleIndex = sheet.OleObjects.Add(5, 2, 200, 300, imageData);
                OleObject oleObject = sheet.OleObjects[oleIndex];

                // Preserve original proportions during resizing by locking the aspect ratio
                oleObject.IsAspectRatioLocked = true;

                // Verify the property is set
                Console.WriteLine("IsAspectRatioLocked: " + oleObject.IsAspectRatioLocked);
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. OleObject will not be added.");
            }

            // Save the workbook
            string outputPath = "OleObjectAspectRatioDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
