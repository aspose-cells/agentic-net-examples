using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class InsertBackgroundImageDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the background image file (adjust as needed)
            string imagePath = "background.jpg";

            if (File.Exists(imagePath))
            {
                // Load the image file into a byte array
                byte[] imageData = File.ReadAllBytes(imagePath);

                // Set the worksheet background image
                worksheet.BackgroundImage = imageData;
                Console.WriteLine("Background image applied.");
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping background image.");
            }

            // Save the workbook
            workbook.Save("WorksheetWithBackground.xlsx");

            Console.WriteLine("Workbook saved successfully.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            InsertBackgroundImageDemo.Run();
        }
    }
}