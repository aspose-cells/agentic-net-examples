using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBackgroundDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // Add a few worksheets for demonstration
            workbook.Worksheets[0].Name = "SheetA";
            workbook.Worksheets.Add("SheetB");
            workbook.Worksheets.Add("SheetC");
            workbook.Worksheets.Add("SheetD");

            // ---------- Prepare a collection of background images ----------
            // Assume image files are located in a folder named "Images" next to the executable
            string imagesFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images");
            string[] imageFiles = new string[]
            {
                Path.Combine(imagesFolder, "bg0.jpg"),
                Path.Combine(imagesFolder, "bg1.jpg"),
                Path.Combine(imagesFolder, "bg2.jpg")
            };

            // Load each image file into a byte array
            byte[][] backgroundImages = new byte[imageFiles.Length][];
            for (int i = 0; i < imageFiles.Length; i++)
            {
                if (!File.Exists(imageFiles[i]))
                {
                    Console.WriteLine($"Image file not found: {imageFiles[i]}");
                    return;
                }
                backgroundImages[i] = File.ReadAllBytes(imageFiles[i]);
            }

            // ---------- Apply background images based on worksheet index ----------
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use the worksheet's Index property to select an image
                int imageIndex = sheet.Index % backgroundImages.Length;
                sheet.BackgroundImage = backgroundImages[imageIndex];
            }

            // ---------- Save the workbook ----------
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WorkbookWithBackgrounds.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
    }
}