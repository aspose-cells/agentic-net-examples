using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

public class InsertHeaderImageExample
{
    public static void Run()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a simple 1x1 pixel PNG image in memory
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=";
        byte[] imageData = Convert.FromBase64String(base64Png);

        // Set the picture in the center section of the header (section index 1)
        Picture headerPicture = worksheet.PageSetup.SetHeaderPicture(1, imageData);

        // Insert the picture placeholder into the header script
        worksheet.PageSetup.SetHeader(1, "&G");

        // Optional: adjust picture properties (e.g., scaling)
        // headerPicture.ScaleWidth = 80;   // percentage
        // headerPicture.ScaleHeight = 80;  // percentage

        // Save the workbook
        workbook.Save("HeaderImageDemo.xlsx");
        Console.WriteLine("Workbook saved with header image.");
    }
}

public class Program
{
    public static void Main()
    {
        InsertHeaderImageExample.Run();
    }
}