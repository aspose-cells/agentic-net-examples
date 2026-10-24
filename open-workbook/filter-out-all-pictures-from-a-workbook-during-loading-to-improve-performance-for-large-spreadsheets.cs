// Title: Remove all pictures from an Excel workbook using Aspose.Cells for .NET to speed up loading of large files
// Common Searches: Aspose.Cells C# remove all images from workbook before saving | how to clear pictures collection in each worksheet with Aspose.Cells | optimize loading of large Excel files by stripping pictures using Aspose.Cells .NET | C# code to delete all embedded pictures from an .xlsx workbook with Aspose.Cells | performance improvement by removing pictures from Excel workbook Aspose.Cells
// Tags: clear worksheet pictures Aspose.Cells | strip images from Excel workbook .NET | optimize large spreadsheet loading Aspose.Cells | remove embedded pictures C# Aspose.Cells | picture removal performance Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the source file, loads the workbook with Aspose.Cells, iterates through each worksheet to clear its Pictures collection, saves the workbook without images, and handles any exceptions, providing a fast way to load large Excel files without picture overhead.
class Program
{
    static void Main()
    {
        const string inputPath = "LargeWorkbook.xlsx";
        const string outputPath = "LargeWorkbook_NoPictures.xlsx";

        try
        {
            // Verify that the source workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook normally.
            Workbook workbook = new Workbook(inputPath);

            // Remove all pictures from each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Pictures.Clear();
            }

            // Save the workbook without pictures.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved without pictures to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
