using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorksheetBackgroundImageDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Simple 1x1 PNG image (transparent) as byte array
            byte[] imageData = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

            // Set the background image for the worksheet (appears behind all content)
            worksheet.BackgroundImage = imageData;

            // Save the workbook with the background image applied
            workbook.Save("WorksheetWithBackground.xlsx");

            Console.WriteLine("Worksheet background image set successfully.");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            WorksheetBackgroundImageDemo.Run();
        }
    }
}