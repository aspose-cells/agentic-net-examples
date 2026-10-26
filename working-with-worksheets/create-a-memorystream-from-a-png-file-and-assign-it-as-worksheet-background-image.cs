// Title: How to assign a PNG image as a worksheet background using a MemoryStream in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet background image from a MemoryStream | load PNG into Excel sheet background using Aspose.Cells .NET | assign background image bytes to a worksheet with Aspose.Cells | create workbook and set sheet background from PNG file in C# | use Aspose.Cells to add PNG background to first worksheet
// Tags: BackgroundImage property PNG Aspose.Cells | MemoryStream worksheet background .NET | load PNG bytes into Aspose.Cells workbook | assign image byte array to worksheet C# | add PNG as worksheet background Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example reads a PNG file into a byte array, creates a new Workbook, retrieves the first worksheet, sets the PNG bytes to the worksheet's BackgroundImage property, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            const string backgroundPath = "background.png";
            const string outputPath = "output.xlsx";

            // Verify that the background image file exists
            if (!File.Exists(backgroundPath))
            {
                Console.WriteLine($"Error: File \"{backgroundPath}\" not found.");
                return;
            }

            // Load the PNG file into a byte array
            byte[] pngBytes = File.ReadAllBytes(backgroundPath);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Assign the image bytes as the worksheet background image
            sheet.BackgroundImage = pngBytes;

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
