using System;
using System.IO;
using Aspose.Cells;

class HtmlToPdfWithIcc
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";
            const string iccProfilePath = @"C:\Profiles\CustomProfile.icc";

            // Verify that the input HTML file exists.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the HTML file into a workbook.
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // If a custom ICC profile is available, embed it (property exists in recent versions).
            if (File.Exists(iccProfilePath))
            {
                // Uncomment the following line if your Aspose.Cells version supports ColorProfile.
                // pdfOptions.ColorProfile = iccProfilePath;
            }

            // Save the workbook as a PDF file with the specified options.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully created at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}