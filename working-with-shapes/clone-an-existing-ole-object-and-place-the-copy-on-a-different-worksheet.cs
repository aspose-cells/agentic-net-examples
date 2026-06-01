using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class CloneOleObjectDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet (source)
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Load snapshot image for the OLE object's preview (use empty array if missing)
            string imagePath = "sampleImage.png";
            byte[] snapshot = File.Exists(imagePath) ? File.ReadAllBytes(imagePath) : new byte[0];

            // Add an OLE object to the source worksheet
            int oleIndex = sourceSheet.OleObjects.Add(5, 2, 200, 300, snapshot);
            OleObject sourceOle = sourceSheet.OleObjects[oleIndex];

            // Load embedded OLE data (use empty array if missing)
            string dataPath = "sampleData.xlsx";
            byte[] embeddedData = File.Exists(dataPath) ? File.ReadAllBytes(dataPath) : new byte[0];
            sourceOle.ObjectData = embeddedData;

            // Add a new worksheet where the clone will be placed
            Worksheet targetSheet = workbook.Worksheets.Add("Target");

            // Clone the OLE object to the target sheet
            int clonedIndex = targetSheet.OleObjects.Add(
                sourceOle.UpperLeftRow,
                sourceOle.UpperLeftColumn,
                sourceOle.Height,
                sourceOle.Width,
                sourceOle.ImageData);

            OleObject clonedOle = targetSheet.OleObjects[clonedIndex];
            clonedOle.ObjectData = sourceOle.ObjectData;               // embedded file data
            clonedOle.FileFormatType = sourceOle.FileFormatType;       // file format
            clonedOle.DisplayAsIcon = sourceOle.DisplayAsIcon;         // display mode
            clonedOle.ProgID = sourceOle.ProgID;                       // program identifier

            // Save the workbook with the cloned OLE object
            workbook.Save("ClonedOleObject.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}