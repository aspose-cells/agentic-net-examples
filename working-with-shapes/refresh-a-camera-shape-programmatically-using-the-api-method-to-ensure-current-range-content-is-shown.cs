using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class RefreshCameraShapeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Initial data in the range that will be captured by the camera
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue("Value 1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Value 2");
                sheet.Cells["B3"].PutValue(20);

                // Add a camera picture that captures the range A1:B3
                // The picture will be placed with its top‑left corner at row 5, column 1
                PictureCollection pictures = sheet.Pictures;
                int pictureIndex = pictures.Camera(5, 1, "A1:B3");

                // Modify the source range to simulate data change
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);

                // Refresh the camera picture: remove the old picture and create a new one
                if (pictureIndex >= 0 && pictureIndex < pictures.Count)
                {
                    pictures.RemoveAt(pictureIndex);
                }
                pictureIndex = pictures.Camera(5, 1, "A1:B3");

                // Save the workbook with the refreshed camera picture
                string outputPath = "RefreshCameraShapeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RefreshCameraShapeDemo.Run();
        }
    }
}