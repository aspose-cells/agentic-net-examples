using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleLabelDemo
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

                // Dummy OLE object data (a simple text file)
                byte[] oleData = Encoding.UTF8.GetBytes("Demo OLE content");

                // Add an OLE object to the worksheet
                int oleIndex = worksheet.OleObjects.Add(5, 5, 100, 100, oleData);
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Assign a custom display label to the OLE object
                oleObject.Label = "MyCustomLabel";

                // Save the workbook
                string outputPath = "OleLabelDemo.xlsx";
                workbook.Save(outputPath);

                // Load the workbook back to verify the label
                if (File.Exists(outputPath))
                {
                    Workbook loadedWorkbook = new Workbook(outputPath);
                    OleObject loadedOleObject = loadedWorkbook.Worksheets[0].OleObjects[0];
                    Console.WriteLine("OLE Object Label: " + loadedOleObject.Label);
                }
                else
                {
                    Console.WriteLine($"File not found: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}