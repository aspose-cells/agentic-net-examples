using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectAspectRatioDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Path to the image that will be used for the OLE object
                string imagePath = "sample.jpg";
                byte[] imageData = null;

                // Load image data only if the file exists
                if (File.Exists(imagePath))
                {
                    imageData = File.ReadAllBytes(imagePath);
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. OLE object will not be added.");
                }

                // Add OLE object only when image data is available
                if (imageData != null && imageData.Length > 0)
                {
                    // Add OLE object at row 5, column 2 with size 200x150 pixels
                    int oleIndex = sheet.OleObjects.Add(5, 2, 200, 150, imageData);
                    OleObject oleObject = sheet.OleObjects[oleIndex];

                    // Lock aspect ratio to preserve original proportions
                    oleObject.IsAspectRatioLocked = true;
                }

                // Save the workbook
                workbook.Save("OleObjectAspectRatioLocked.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}