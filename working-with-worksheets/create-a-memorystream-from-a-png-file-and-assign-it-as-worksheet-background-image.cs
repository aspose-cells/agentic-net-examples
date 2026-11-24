// Title: How to set a PNG file as a worksheet background image using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet background image from local PNG file | Load PNG into byte array and apply as Excel sheet background with Aspose.Cells | Assign background picture to first worksheet using Aspose.Cells workbook API | C# code to embed PNG as worksheet background in generated Excel file | Aspose.Cells background image property example for .NET developers
// Tags: worksheet background image Aspose.Cells C# | load PNG byte array for Excel background | Aspose.Cells set background picture from file | Excel workbook background image using Aspose.Cells | C# embed PNG as worksheet background

using System;
using System.IO;
using Aspose.Cells;

// The example checks for a PNG file, reads it into a byte array, creates a new Workbook, accesses the first worksheet, assigns the PNG bytes to the worksheet's BackgroundImage property, and saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            const string imagePath = "image.png";
            const string outputPath = "output.xlsx";

            // Verify that the PNG file exists before attempting to read it
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Error: The image file \"{imagePath}\" was not found.");
                return;
            }

            // Load the PNG file into a byte array
            byte[] pngBytes = File.ReadAllBytes(imagePath);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Assign the PNG bytes as the worksheet background image
            worksheet.BackgroundImage = pngBytes;

            // Save the workbook to an Excel file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
