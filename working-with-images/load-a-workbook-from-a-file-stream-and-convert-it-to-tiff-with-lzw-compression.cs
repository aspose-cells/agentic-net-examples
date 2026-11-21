// Title: Load an Excel workbook from a FileStream and save it as a TIFF image with Aspose.Cells for .NET
// Common Searches: how to load an Excel workbook from a FileStream and export to TIFF using Aspose.Cells in C# | Aspose.Cells C# save workbook as TIFF image directly from a stream | convert XLSX to TIFF without intermediate file using Aspose.Cells | C# Aspose.Cells TIFF export options and compression settings | read Excel file from memory stream and generate TIFF with Aspose.Cells
// Tags: load workbook from filestream Aspose.Cells | save workbook as tiff Aspose.Cells | excel to tiff conversion C# | aspnet cells image export tiff | tiff export compression options Aspose.Cells | stream based workbook loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook via a FileStream, and saves the workbook as a TIFF image using the SaveFormat.Tiff option, which applies the default compression.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.tiff";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook;
            using (FileStream inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                workbook = new Workbook(inputStream);
            }

            // Save the workbook as a TIFF image (default compression)
            workbook.Save(outputPath, SaveFormat.Tiff);
            Console.WriteLine($"Workbook successfully saved as TIFF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
