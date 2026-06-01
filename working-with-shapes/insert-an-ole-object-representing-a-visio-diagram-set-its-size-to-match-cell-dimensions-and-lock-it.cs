using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

class InsertVisioOleObject
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Load the Visio file (the actual OLE data) into a byte array
            byte[] visioData = File.Exists("diagram.vsdx")
                ? File.ReadAllBytes("diagram.vsdx")
                : Array.Empty<byte>(); // fallback to empty if file not found

            // Load an image that will be shown as the OLE object's preview (optional)
            byte[] previewImage = File.Exists("visio_icon.png")
                ? File.ReadAllBytes("visio_icon.png")
                : Array.Empty<byte>(); // fallback to empty if file not found

            // Define the cell where the OLE object will be placed (e.g., B2 -> row 1, column 1)
            int topRow = 1;
            int leftColumn = 1;

            // Approximate the cell size in pixels
            double rowHeightPoints = worksheet.Cells.Rows[topRow].Height; // default 12.75 points
            int heightPixels = (int)Math.Round(rowHeightPoints * 96 / 72);

            double columnWidthChars = worksheet.Cells.Columns[leftColumn].Width; // default 8.43 characters
            int widthPixels = (int)Math.Round(columnWidthChars * 7);

            // Add the OLE object with the preview image (if previewImage is empty, Excel will use default icon)
            int oleIndex = worksheet.OleObjects.Add(topRow, leftColumn, heightPixels, widthPixels, previewImage);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Embed the Visio file data into the OLE object (if visioData is empty, the object will be created without embedded data)
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: visioData,
                sourceFileName: "diagram.vsdx",
                displayAsIcon: false,
                label: "Visio Diagram");

            // Set the ProgID for Visio so Excel knows which application to use
            oleObject.ProgID = "Visio.Drawing";

            // Lock the object (prevents moving/resizing when the sheet is protected)
            oleObject.IsLocked = true;

            // Make the object move and size with its underlying cells
            oleObject.Placement = PlacementType.MoveAndSize;

            // Save the workbook
            workbook.Save("VisioOleObject.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}