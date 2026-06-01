using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsOleObjectExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths
                string workbookPath = "input.xlsx";
                string csvFilePath = "data.csv";

                // Optional icon for the OLE object; null if the file does not exist
                byte[] iconData = File.Exists("icon.png") ? File.ReadAllBytes("icon.png") : null;

                // Ensure the CSV file exists (required for the linked OLE object)
                if (!File.Exists(csvFilePath))
                {
                    throw new FileNotFoundException("CSV file not found.", csvFilePath);
                }

                // Load existing workbook or create a new one if the file is missing
                Workbook workbook = File.Exists(workbookPath) ? new Workbook(workbookPath) : new Workbook();

                Worksheet sheet = workbook.Worksheets[0];

                // Add a linked OLE object that points to the CSV file
                // Parameters: topRow, leftColumn, height (pixels), width (pixels), imageData, linkedFile
                int oleIndex = sheet.OleObjects.Add(2, 2, 150, 200, iconData, csvFilePath);

                // Retrieve the newly added OleObject
                OleObject ole = sheet.OleObjects[oleIndex];

                // Set the program ID so the CSV opens with Excel
                ole.ProgID = "Excel.CSV";

                // Load automatically when the workbook is opened
                ole.AutoLoad = true;

                // Show the OLE object as an icon
                ole.DisplayAsIcon = true;

                // Save the modified workbook
                workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}