// Title: Add a branding PNG image to the first worksheet of an existing Excel file using Aspose.Cells for .NET
// AI Prompts: Add a PNG logo to cell A1 of the first worksheet in a workbook loaded with Aspose.Cells for .NET. | Create a reusable method that takes a workbook, worksheet index, and image path, then inserts the image using Aspose.Cells. | Modify the sample to apply the image as a worksheet background instead of a floating picture with Aspose.Cells.
// Common Searches: how to add a logo to the first sheet of an Excel workbook using Aspose.Cells in C# | example of using worksheet.Pictures.Add to place an image at cell A1 | changing a worksheet's background to a PNG file with Aspose.Cells .NET
// Tags: Aspose.Cells picture insertion API | C# embed PNG into Excel worksheet | Excel branding with Aspose.Cells | load workbook and insert image Aspose.Cells | set worksheet background image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that Input.xlsx and BrandImage.png exist, loads the workbook, accesses the first worksheet, inserts the PNG image at the top‑left corner using worksheet.Pictures.Add, and saves the result as Output.xlsx, handling any exceptions that occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string imagePath = "BrandImage.png";
            const string outputPath = "Output.xlsx";

            // Ensure required files exist
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Background image not found: {imagePath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);
            var worksheet = workbook.Worksheets[0];

            // Insert the image onto the worksheet (top‑left corner)
            worksheet.Pictures.Add(0, 0, imagePath);

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
