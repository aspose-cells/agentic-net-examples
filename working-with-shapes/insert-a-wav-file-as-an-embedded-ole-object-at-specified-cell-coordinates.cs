using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsOleObjectDemo
{
    public class InsertWavAsOleObject
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the position (row, column) and size (height, width) for the OLE object
            int topRow = 5;          // Upper left row index (zero‑based)
            int leftColumn = 2;      // Upper left column index (zero‑based)
            int height = 100;        // Height in pixels
            int width = 100;         // Width in pixels

            // Load an image that will be shown as the OLE object's icon
            const string iconPath = "icon.png";
            if (!File.Exists(iconPath))
                throw new FileNotFoundException($"Icon file not found: {iconPath}");
            byte[] iconData = File.ReadAllBytes(iconPath);

            // Add a placeholder OLE object to the worksheet using the icon image
            int oleIndex = worksheet.OleObjects.Add(topRow, leftColumn, height, width, iconData);
            OleObject oleObject = worksheet.OleObjects[oleIndex];

            // Load the WAV file that will be embedded
            const string wavPath = "audio.wav";
            if (!File.Exists(wavPath))
                throw new FileNotFoundException($"WAV file not found: {wavPath}");
            byte[] wavData = File.ReadAllBytes(wavPath);

            // Embed the WAV file into the OLE object
            oleObject.SetEmbeddedObject(
                linkToFile: false,
                objectData: wavData,
                sourceFileName: wavPath,
                displayAsIcon: true,
                label: "Audio");

            // Optional: set additional properties
            oleObject.DisplayAsIcon = true;   // Show as an icon
            oleObject.AutoLoad = false;      // Do not auto‑load when workbook opens

            // Save the workbook
            const string outputPath = "WorkbookWithWavOleObject.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}