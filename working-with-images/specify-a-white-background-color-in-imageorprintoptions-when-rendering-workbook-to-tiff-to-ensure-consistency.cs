// Title: How to set a white background with ImageOrPrintOptions when saving a workbook as TIFF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set white background for TIFF export using ImageOrPrintOptions | Specify background color when rendering Excel to TIFF in .NET | ImageOrPrintOptions background color property for TIFF output Aspose.Cells | Save workbook to TIFF with consistent white background in C# | C# Aspose.Cells render workbook to TIFF with custom background color
// Tags: ImageOrPrintOptions background color TIFF | Aspose.Cells set TIFF background white | C# render workbook to TIFF with white background | Aspose.Cells TIFF export background configuration | Excel to TIFF conversion white background .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, configures ImageOrPrintOptions to use a white background, and saves the workbook as a TIFF image, ensuring consistent background color with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as a TIFF file
            workbook.Save(outputPath, SaveFormat.Tiff);
            Console.WriteLine($"Workbook successfully saved as TIFF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
