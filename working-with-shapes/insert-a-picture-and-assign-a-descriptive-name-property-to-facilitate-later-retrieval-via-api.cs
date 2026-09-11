// Title: Insert a PNG image into an Excel worksheet at cell A1 and set a custom Name property with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to add a picture from a PNG file to cell A1 of a new workbook and assign the picture's Name property to "CompanyLogo". | Create a workbook, insert a PNG image at the top‑left cell, give the picture a descriptive identifier, and save the workbook as XLSX using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to add a picture to a specific cell and name it | set custom name for inserted image in Excel using Aspose.Cells .NET | retrieve picture by Name property in Aspose.Cells workbook | insert PNG logo into Excel worksheet with Aspose.Cells and assign identifier
// Tags: add picture to worksheet Aspose.Cells C# | picture Name attribute Aspose.Cells | insert PNG into XLSX Aspose.Cells | custom identifier for Excel picture Aspose.Cells | retrieve picture by Name Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, inserts a PNG image at cell A1, assigns the picture's Name property the value "CompanyLogo", and saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = "sample.png"; // Replace with your actual image file path

            // Ensure the image file exists before attempting to insert
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Insert the picture at cell A1 (row 0, column 0)
            // The Add method returns the index of the newly added picture
            int pictureIndex = sheet.Pictures.Add(0, 0, imagePath);

            // Retrieve the picture object using the returned index
            Picture picture = sheet.Pictures[pictureIndex];

            // Assign a descriptive name to the picture for later retrieval via API
            picture.Name = "CompanyLogo";

            // Save the workbook (lifecycle rule: save)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
