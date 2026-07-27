using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class IncreasePictureBrightness
{
    static void Main()
    {
        // Load an existing workbook that contains a PNG picture
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one picture on the worksheet
        if (sheet.Pictures.Count == 0)
        {
            Console.WriteLine("No pictures found on the worksheet.");
            return;
        }

        // Get the first picture (assumed to be a PNG image)
        Picture picture = sheet.Pictures[0];

        // Increase brightness by 20%
        // The Brightness property range is -100 to 100 (percentage)
        // Adding 20 to the current value raises the brightness by 20%
        double currentBrightness = picture.FormatPicture.Brightness;
        double newBrightness = currentBrightness + 20.0;

        // Clamp the value to the allowed range
        if (newBrightness > 100.0) newBrightness = 100.0;
        if (newBrightness < -100.0) newBrightness = -100.0;

        picture.FormatPicture.Brightness = newBrightness;

        Console.WriteLine($"Brightness changed from {currentBrightness} to {newBrightness}.");

        // Save the workbook with the updated picture brightness
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}