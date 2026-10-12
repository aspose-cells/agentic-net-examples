// Title: Insert a PNG picture into cell B2, set a custom Name, and resize it using Aspose.Cells for .NET
// AI Prompts: Insert a PNG file into cell B2 of a new workbook, assign the picture's Name property to "CompanyLogo", set its width to 150 px and height to 100 px, then save as Xlsx with Aspose.Cells C#. | Create a picture shape on a worksheet, give it a descriptive name for later retrieval, adjust its dimensions, and persist the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# insert image into specific cell and name it | How to set the Name property of a picture shape in Aspose.Cells | Resize picture after adding to worksheet with Aspose.Cells .NET | Retrieve picture by custom name using Aspose.Cells API | Save workbook with embedded PNG image using Aspose.Cells
// Tags: insert picture into worksheet cell Aspose.Cells | assign custom name to picture shape Aspose.Cells | resize picture dimensions Aspose.Cells C# | save workbook with embedded PNG Aspose.Cells | retrieve picture by name Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a PNG image to cell B2, sets the picture's Name property to "CompanyLogo", adjusts its width and height, and saves the file as WorkbookWithPicture.xlsx.
class InsertPictureExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the image file to be inserted
            string imagePath = "sample_image.png";

            // Ensure the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Add the picture to the worksheet at cell B2 (row 1, column 1)
            using (FileStream imageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                // rowIndex = 1 (second row), columnIndex = 1 (second column)
                int pictureIndex = sheet.Pictures.Add(1, 1, imageStream);
                Picture picture = sheet.Pictures[pictureIndex];

                // Assign a descriptive name to the picture for later retrieval
                picture.Name = "CompanyLogo";

                // Optionally adjust picture size (width and height in pixels)
                picture.Width = 150;
                picture.Height = 100;
            }

            // Save the workbook to a file
            string outputPath = "WorkbookWithPicture.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
