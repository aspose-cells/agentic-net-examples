// Title: Load an Excel workbook from a FileStream and save it as a TIFF image with Aspose.Cells for .NET
// AI Prompts: Write a C# program that opens an .xlsx file using a FileStream, loads it into an Aspose.Cells Workbook, and saves the workbook as a TIFF image. | Create a C# example that checks whether an Excel file exists, reads it via FileStream into Aspose.Cells, and exports the workbook to a TIFF file.
// Common Searches: c# asp.net load excel file from filestream and export to tiff using aspose.cells | how to convert xlsx to tiff image with aspose.cells in .net core | asp.net core save workbook as tiff image from file stream example | aspose.cells c# save workbook as tiff image with error handling
// Tags: Aspose.Cells load workbook from FileStream | Aspose.Cells export workbook to TIFF | C# convert Excel to TIFF image | validate Excel file existence before conversion | Aspose.Cells TIFF save options .NET

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing.Imaging; // Required for ImageFormat

// The sample checks that the source Excel file exists, opens it through a FileStream, loads it into an Aspose.Cells Workbook, and then saves the workbook as a TIFF image using the SaveFormat.Tiff option, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.tiff";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from a file stream
            using (FileStream inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                Workbook workbook = new Workbook(inputStream);

                // Save the workbook as a TIFF file
                workbook.Save(outputPath, SaveFormat.Tiff);
                Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
