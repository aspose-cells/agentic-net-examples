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

            // Paths to the image representing the OLE object and the file to embed
            string imagePath = "oleImage.jpg";
            string oleFilePath = "embedded.xlsx";

            // Load image data; if missing, use a 1x1 PNG placeholder
            byte[] imageData;
            if (File.Exists(imagePath))
            {
                imageData = File.ReadAllBytes(imagePath);
            }
            else
            {
                // Base64‑encoded 1x1 transparent PNG
                const string placeholderBase64 =
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2V8AAAAASUVORK5CYII=";
                imageData = Convert.FromBase64String(placeholderBase64);
            }

            // Load OLE file data if it exists; otherwise use an empty array
            byte[] oleData = File.Exists(oleFilePath) ? File.ReadAllBytes(oleFilePath) : Array.Empty<byte>();

            // Define OLE object position (row, column) and size (pixels)
            int topRow = 5;
            int leftColumn = 2;
            int height = 200;
            int width = 300;

            // Add the OLE object to the worksheet using the image as its preview
            int oleIndex = sheet.OleObjects.Add(topRow, leftColumn, height, width, imageData);
            OleObject ole = sheet.OleObjects[oleIndex];

            // Embed the actual file data into the OLE object if any data is present
            if (oleData.Length > 0)
            {
                ole.ObjectData = oleData;
            }
            ole.DisplayAsIcon = false; // show preview image instead of an icon

            // Freeze rows above and columns left of the OLE object
            sheet.FreezePanes(topRow, leftColumn, topRow, leftColumn);

            // Save the workbook
            workbook.Save("OleObjectWithFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}