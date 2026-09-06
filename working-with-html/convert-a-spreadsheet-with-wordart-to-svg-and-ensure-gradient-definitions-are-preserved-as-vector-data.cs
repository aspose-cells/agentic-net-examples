// Title: Convert Excel worksheets containing WordArt to SVG while preserving gradient fills as vector graphics using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with WordArt and saves each worksheet as an SVG, ensuring gradient fills are kept as vector elements. | Show how to configure ImageSaveOptions for SVG in Aspose.Cells to enable OnePagePerSheet and retain WordArt gradient definitions. | Adapt the example to loop through all worksheets and output separate SVG files while preserving all WordArt styling.
// Common Searches: Aspose.Cells export WordArt to SVG preserving gradient fill .NET | How to keep Excel WordArt gradients when saving as SVG with C# | SVG output from Aspose.Cells losing gradient definitions | OnePagePerSheet option for SVG conversion in Aspose.Cells | Convert multiple Excel sheets with WordArt to separate SVG files using Aspose.Cells
// Tags: Aspose.Cells ImageSaveOptions SVG export | WordArt gradient vector retention | OnePagePerSheet setting for SVG | Excel to SVG conversion .NET | gradient fill preservation in SVG conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program verifies the input .xlsx file, loads it into an Aspose.Cells Workbook, configures ImageSaveOptions with SaveFormat.Svg and OnePagePerSheet enabled, and saves the workbook as an SVG file while handling any runtime exceptions.
class WordArtToSvgConverter
{
    static void Main()
    {
        // Path to the source Excel file containing WordArt
        string inputFile = "input.xlsx";

        // Path where the resulting SVG will be saved
        string outputFile = "output.svg";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure SVG save options using the recommended ImageSaveOptions API
            ImageSaveOptions svgOptions = new ImageSaveOptions(SaveFormat.Svg)
            {
                // Ensure each sheet is saved as a single page (required for proper SVG output)
                ImageOrPrintOptions = { OnePagePerSheet = true }
            };

            // Save the workbook as SVG with the specified options
            workbook.Save(outputFile, svgOptions);

            Console.WriteLine($"SVG file saved successfully to: {outputFile}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, saving)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
