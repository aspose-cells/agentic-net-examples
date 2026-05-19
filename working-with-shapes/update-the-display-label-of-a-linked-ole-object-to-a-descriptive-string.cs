using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectLabelUpdateDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Minimal 1x1 PNG image (transparent) used as placeholder for the OLE object's icon
                byte[] placeholderImage = new byte[]
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

                // Add an OLE object to the worksheet using the placeholder image
                int oleIndex = worksheet.OleObjects.Add(10, 10, 200, 200, placeholderImage);
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Update the display label of the OLE object
                oleObject.Label = "Descriptive OLE Object Label";

                // Save the workbook
                string filePath = "OleObjectLabelDemo.xlsx";
                workbook.Save(filePath);

                // Verify the label after reloading the workbook
                if (File.Exists(filePath))
                {
                    Workbook loadedWorkbook = new Workbook(filePath);
                    OleObject loadedOleObject = loadedWorkbook.Worksheets[0].OleObjects[0];
                    Console.WriteLine("OLE Object Label after reload: " + loadedOleObject.Label);
                }
                else
                {
                    Console.WriteLine("Error: The file was not saved correctly.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}