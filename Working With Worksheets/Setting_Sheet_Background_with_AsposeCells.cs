using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Load background image from file; if not found, use a simple 1x1 PNG
        byte[] imageData;
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "background.jpg");
        if (File.Exists(imagePath))
        {
            imageData = File.ReadAllBytes(imagePath);
        }
        else
        {
            // 1x1 pixel PNG (transparent)
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
            imageData = Convert.FromBase64String(base64Png);
        }

        // Set the worksheet background image
        worksheet.BackgroundImage = imageData;

        // Save the workbook with the background applied
        workbook.Save("WorksheetWithBackground.xlsx");

        Console.WriteLine("Worksheet background image set successfully.");
    }
}