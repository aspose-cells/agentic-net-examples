// Title: Aspose.Cells C# – Lock OleObject Aspect Ratio
// Description: This C# example creates a workbook, inserts an OleObject with a preview image (or a generated placeholder), enables the IsAspectRatioLocked flag so the object retains its original width‑height ratio when resized, and saves the file as OleObjectAspectRatioLocked.xlsx.
// Keywords: Aspose.Cells OleObject aspect ratio | IsAspectRatioLocked C# | lock OleObject proportions | OleObject resizing Aspose | C# embed OleObject | Aspose.Cells shape properties | preserve object ratio | Excel OleObject scaling
// Common Searches: Aspose.Cells lock OleObject aspect ratio | IsAspectRatioLocked property example | keep OleObject proportions C# | prevent distortion OleObject Aspose | set OleObject scaling flag
// Developer Intent: Enable the OleObject's IsAspectRatioLocked flag to keep its original width‑height ratio during any resize operation.
// Use Cases: Embed a chart as an OleObject in a financial report and ensure it scales uniformly when columns are adjusted. | Add a PDF preview to a template while preventing distortion during printing or export. | Generate a workbook with multiple embedded objects that must retain their native aspect ratios across devices. | Create a dynamic dashboard where users can resize rows or columns without stretching embedded images.
// AI Prompts: Write C# code using Aspose.Cells to add an OleObject and set IsAspectRatioLocked to true. | Explain how the IsAspectRatioLocked property influences OleObject scaling compared to manually setting Width and Height. | Show how to generate a fallback PNG placeholder for an OleObject preview when the source image is unavailable. | Demonstrate retrieving and modifying the aspect‑ratio lock of an existing OleObject in a workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, inserts an OleObject with a preview image (or a generated placeholder), enables the IsAspectRatioLocked flag so the object retains its original width‑height ratio when resized, and saves the file as OleObjectAspectRatioLocked.xlsx.
class SetOleObjectAspectRatio
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Load image data for the OleObject preview; use placeholder if file is missing
            string imagePath = "OleObjectPreview.png";
            byte[] imageData = File.Exists(imagePath) ? File.ReadAllBytes(imagePath) : CreatePlaceholderPng();

            // Add an OleObject to the worksheet (row, column, height, width, preview image)
            int oleIndex = sheet.OleObjects.Add(5, 2, 200, 300, imageData);
            OleObject oleObject = sheet.OleObjects[oleIndex];

            // Preserve original proportions during resizing
            oleObject.IsAspectRatioLocked = true;

            // Save the workbook with the configured OleObject
            workbook.Save("OleObjectAspectRatioLocked.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // Generates a minimal 1x1 pixel PNG to use when the preview image is not found
    private static byte[] CreatePlaceholderPng()
    {
        return new byte[]
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
    }
}
