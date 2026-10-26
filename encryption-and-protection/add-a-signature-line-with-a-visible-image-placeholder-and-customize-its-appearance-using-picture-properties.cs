// Title: Add a signature line with a placeholder PNG image and format its size, position, and border using Aspose.Cells for .NET (C#)
// Common Searches: how to insert a placeholder PNG as a signature line in an Excel file using Aspose.Cells C# | Aspose.Cells set picture width height and border thickness for a signature image | position a picture at a specific cell location in a workbook with Aspose.Cells .NET | add and format a signature picture on a worksheet using Aspose.Cells API | save workbook after adding signature line image with Aspose.Cells for .NET
// Tags: insert picture as signature line Aspose.Cells | set picture dimensions Aspose.Cells C# | format picture border Aspose.Cells | position picture on worksheet Aspose.Cells | add placeholder image to Excel workbook Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, checks for a placeholder PNG file, inserts it as a picture at cell A1, customizes its width, height, left, top, rotation angle, and border thickness, and then saves the workbook as SignatureLine.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the placeholder image (ensure the file exists)
            string imagePath = "placeholder.png";

            if (File.Exists(imagePath))
            {
                try
                {
                    // Add the picture to the worksheet at cell A1 (row 0, column 0)
                    int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);
                    Picture pic = sheet.Pictures[pictureIndex];

                    // Set picture size (pixels)
                    pic.Width = 150;
                    pic.Height = 50;

                    // Position the picture (pixels from the top‑left corner of the sheet)
                    pic.Left = 100;
                    pic.Top = 200;

                    // Additional formatting
                    pic.RotationAngle = 0; // No rotation

                    // Set border thickness
                    pic.Line.Weight = 1; // Border thickness

                    // Note: Setting border color is not supported directly via LineFormat in some versions.
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to insert picture: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
            }

            // Save the workbook
            string outputPath = "SignatureLine.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
