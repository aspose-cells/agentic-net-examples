// Title: C# – Hide an OLE object in Excel with Aspose.Cells (OleObject.IsHidden)
// Description: Shows how to create a workbook, add a Word‑document OLE object with a placeholder PNG, set its ProgID, hide it using the OleObject.IsHidden property, and save the file.
// Keywords: Aspose.Cells C# hide OLE object | OleObject.IsHidden property | Excel OLE visibility .NET | embed Word document as OLE | background OLE objects Aspose.Cells | Excel hidden objects C# | Aspose.Cells example hide OleObject | set OleObject visibility false
// Common Searches: Aspose.Cells hide OLE object C# example | OleObject.IsHidden true Aspose.Cells | make embedded Word document invisible Excel | how to hide OLE objects in generated spreadsheet | C# Aspose.Cells set OLE object visibility
// Developer Intent: Hide an OLE object so it does not appear on the worksheet.
// Use Cases: Store a Word document as a hidden OLE object for later extraction. | Add a preview image for an OLE object then hide it to keep the sheet tidy. | Create a template where OLE objects start hidden and are revealed by a separate process.
// AI Prompts: Generate C# code with Aspose.Cells that inserts an OLE object and sets its visibility to false. | Retrieve an existing OleObject from a worksheet and hide it using the appropriate property. | Explain the difference between OleObject.IsHidden and OleObject.Visible in Aspose.Cells and when to use each.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectVisibilityDemo
{
    // Shows how to create a workbook, add a Word‑document OLE object with a placeholder PNG, set its ProgID, hide it using the OleObject.IsHidden property, and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Generate a simple PNG image for OLE object preview
                byte[] imageData = GeneratePlaceholderImage();

                // Add an OLE object at row 5, column 2 with size 150x150 pixels
                int oleIndex = sheet.OleObjects.Add(5, 2, 150, 150, imageData);

                // Retrieve the added OLE object
                OleObject ole = sheet.OleObjects[oleIndex];

                // Set the OLE object's ProgID (embed a Word document)
                ole.ProgID = "Word.Document";

                // Hide the OLE object
                ole.IsHidden = true;

                // Save the workbook
                string outputPath = "OleObjectHiddenDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to return a 1x1 pixel PNG image byte array
        private static byte[] GeneratePlaceholderImage()
        {
            // Base64-encoded 1x1 white PNG
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            return Convert.FromBase64String(base64Png);
        }
    }
}
