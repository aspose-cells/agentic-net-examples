using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadXlsxWithEmbeddedImagesDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define image files of supported formats
            string[] imageFiles = new string[]
            {
                "sample.png",
                "sample.jpg",
                "sample.bmp",
                "sample.gif"
            };

            // Add each image to the worksheet
            int row = 0;
            foreach (string imgPath in imageFiles)
            {
                if (!File.Exists(imgPath))
                {
                    Console.WriteLine($"Image file not found: {imgPath}");
                    continue;
                }

                // Add picture at the current row, first column (0-indexed)
                sheet.Pictures.Add(row, 0, imgPath);
                row++;
            }

            // Save the workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}