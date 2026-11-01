// Title: Convert Excel workbook to PDF while stripping all pictures using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF without embedded images | How to remove all pictures from a workbook before PDF conversion in .NET | C# clear worksheet Pictures collection to reduce PDF size with Aspose.Cells | Generate minimal size PDF from Excel by excluding linked images using Aspose.Cells | Aspose.Cells remove external linked images from Excel before saving as PDF
// Tags: worksheet pictures removal Aspose.Cells | excel to pdf conversion minimal size | clear pictures collection C# Aspose.Cells | exclude embedded images PDF export | optimize PDF output Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The C# program loads an Excel file with Aspose.Cells, iterates through each worksheet to clear its Pictures collection, and then saves the workbook as a PDF. The resulting PDF omits all embedded or linked images, yielding a smaller file size.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and remove all pictures (external linked images cannot be distinguished via API)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Remove pictures safely
                    if (sheet.Pictures.Count > 0)
                    {
                        // Clear the collection to avoid index shifting issues
                        sheet.Pictures.Clear();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not remove pictures from sheet \"{sheet.Name}\": {ex.Message}");
                }
            }

            // Save the workbook as PDF without the pictures
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
