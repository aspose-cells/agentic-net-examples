using System;
using System.IO;
using Aspose.Cells;

class ColorScaleHtmlExport
{
    static void Main()
    {
        try
        {
            // Path to the existing workbook that already contains a ColorScale conditional formatting rule
            string inputFile = "ColorScaleWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Preserve all formatting details, including conditional formatting such as ColorScale
            htmlOptions.ExcludeUnusedStyles = false;

            // Define the output HTML file path
            string outputFile = "ColorScaleOutput.html";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML with the specified options
            workbook.Save(outputFile, htmlOptions);

            Console.WriteLine("HTML file with ColorScale export saved to: " + outputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}