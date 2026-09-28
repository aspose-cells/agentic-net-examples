// Title: Use Aspose.Cells for .NET to set a PNG worksheet background from a MemoryStream in C#
// AI Prompts: Read a PNG file into a MemoryStream, assign its byte array to Worksheet.BackgroundImage, and save the workbook with Aspose.Cells in C#. | Create a new workbook, apply a PNG background to the first worksheet using a byte array without persisting the image file, and handle missing‑file exceptions. | Programmatically add a PNG picture as the worksheet background and export the Excel file while ensuring the output directory exists with Aspose.Cells.
// Common Searches: aspnet set worksheet background image from memory stream Aspose.Cells | c# Aspose.Cells load png bytes as worksheet background | how to assign background picture to Excel sheet without saving image file using Aspose.Cells | Aspose.Cells background image from byte array example | save workbook with PNG background using Aspose.Cells .NET
// Tags: background image assignment via byte array | Aspose.Cells PNG background integration | C# create workbook with background picture | error handling for missing background file | ensure output directory before saving workbook

using System;
using System.IO;
using Aspose.Cells;

// The sample reads a PNG file, loads its bytes into a MemoryStream, creates a new workbook, sets the first worksheet's BackgroundImage property using the byte array, creates the output folder if needed, and saves the workbook, with proper error handling for missing image files.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the background PNG image
            string pngPath = @"C:\Images\background.png";

            // Verify that the PNG file exists
            if (!File.Exists(pngPath))
                throw new FileNotFoundException("Background image not found.", pngPath);

            // Load PNG file bytes
            byte[] pngBytes = File.ReadAllBytes(pngPath);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Assign the PNG bytes as the worksheet background image
            sheet.BackgroundImage = pngBytes;

            // Ensure the output directory exists
            string outputPath = @"C:\Output\WorkbookWithBackground.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or display the error as needed
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
