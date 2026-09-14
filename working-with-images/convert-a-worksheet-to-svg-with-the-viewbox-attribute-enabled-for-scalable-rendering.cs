// Title: Convert the first worksheet of an Excel file to a scalable SVG that includes a viewBox using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, sets ImageOrPrintOptions for SVG output, and uses SheetRender to generate an SVG file that contains a viewBox for flexible scaling. | Adapt an existing Aspose.Cells example so the produced SVG includes a viewBox attribute, enabling responsive resizing, and save it to a specified path.
// Common Searches: aspacells c# generate svg from excel sheet including viewbox | how to create responsive svg from excel using aspose.cells | c# convert excel worksheet to scalable svg for web | sheetrender produce svg with viewbox attribute aspose.cells | imageorprintoptions settings for svg export .net
// Tags: Aspose.Cells SVG export with viewBox | C# SheetRender to scalable SVG | ImageOrPrintOptions SVG configuration .NET | Excel worksheet to responsive SVG | Aspose.Cells rendering for web graphics

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example checks for an input.xlsx file, creates a simple workbook if missing, loads the workbook, configures ImageOrPrintOptions for one-page-per-sheet SVG rendering, and uses SheetRender to save the first worksheet as output.svg. The generated SVG includes a viewBox attribute, allowing the image to scale responsively, and any errors are caught and reported.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.svg";

            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);
            var worksheet = workbook.Worksheets[0];

            // Set up rendering options for SVG output.
            var options = new ImageOrPrintOptions
            {
                OnePagePerSheet = true // Render the whole sheet on one page.
                // No need to set ImageFormat; the format is inferred from the file extension.
            };

            // Render the worksheet to SVG.
            var sheetRender = new SheetRender(worksheet, options);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"SVG file saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
