using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectNameDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Get a minimal 1x1 PNG image as byte array
                byte[] imageData = GetPlaceholderImage();

                // Add the OLE object to the worksheet
                // Parameters: top row, left column, height (pixels), width (pixels), image data
                int oleIndex = worksheet.OleObjects.Add(5, 2, 200, 150, imageData);
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Assign a unique name using a GUID
                oleObject.Name = "OleObject_" + Guid.NewGuid().ToString("N");

                // Optionally set a label for display purposes
                oleObject.Label = "My Unique OLE Object";

                // Save the workbook
                string outputPath = "OleObjectWithUniqueName.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Returns a 1x1 transparent PNG image as a byte array
        private static byte[] GetPlaceholderImage()
        {
            // Base64-encoded PNG (1x1 pixel, transparent)
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X0ZcAAAAASUVORK5CYII=";
            return Convert.FromBase64String(base64Png);
        }
    }
}