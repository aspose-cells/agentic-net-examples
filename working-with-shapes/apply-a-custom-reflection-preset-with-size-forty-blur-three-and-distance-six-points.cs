// Title: Add a picture to cell C3 and apply a custom reflection (size 40%, blur 3, distance 6 points) using Aspose.Cells for .NET (C#)
// AI Prompts: Insert an image into a worksheet cell and configure its Reflection.Size, Reflection.Blur, and Reflection.Distance properties with Aspose.Cells in C#. | Show how to set a custom reflection effect on a picture shape in an Excel file using the Aspose.Cells .NET API. | Create a new workbook, add a PNG to cell C3, and apply a 40% size, 3‑point blur, and 6‑point distance reflection using Aspose.Cells.
// Common Searches: aspnet c# how to add picture to Excel cell with Aspose.Cells and set reflection properties | set custom reflection size blur distance for picture shape in Aspose.Cells | Aspose.Cells picture reflection example C# | apply reflection effect to image in Excel using Aspose.Cells .NET
// Tags: Aspose.Cells picture reflection configuration | C# add image to Excel cell Aspose.Cells | custom reflection preset Aspose.Cells API | set reflection size blur distance .NET Excel | Excel shape effects Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, inserts a PNG image into cell C3, applies a reflection effect with 40% size, blur radius of 3, and distance of 6 points, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file
            string imagePath = "input.png";

            // Ensure the image file exists before adding it
            if (File.Exists(imagePath))
            {
                // Add the picture to cell C3 (row index 2, column index 2)
                int pictureIndex = sheet.Pictures.Add(2, 2, imagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Apply custom reflection settings
                picture.Reflection.Size = 40;      // size in percentage
                picture.Reflection.Blur = 3;       // blur radius
                picture.Reflection.Distance = 6;  // distance in points
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Save the workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved successfully as output.xlsx.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
