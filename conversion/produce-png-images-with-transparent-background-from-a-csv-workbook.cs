// Title: Create transparent PNG images from each sheet of a CSV workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export CSV worksheet to PNG with transparent background | How to render CSV sheets as transparent PNG files in .NET | Using ImageOrPrintOptions Transparent property to generate PNG from CSV workbook | Convert each worksheet in a CSV file to PNG with alpha channel using Aspose.Cells | C# code to load CSV with LoadOptions and save sheets as transparent PNG
// Tags: CSV workbook to transparent PNG conversion Aspose.Cells | ImageOrPrintOptions PNG transparency setting C# | SheetRender ToImage PNG output Aspose.Cells | LoadOptions LoadFormat.Csv Aspose.Cells | export worksheet as PNG with alpha channel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the CSV file, loads it into an Aspose.Cells Workbook with LoadOptions for CSV, then iterates through all worksheets. For each sheet it configures ImageOrPrintOptions with Transparent enabled, renders the first page via SheetRender, and saves the result as a PNG file (Sheet1.png, Sheet2.png, …) preserving a transparent background, while handling any rendering errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.csv";

            // Verify that the input CSV file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the CSV file into a workbook
            var loadOptions = new LoadOptions(LoadFormat.Csv);
            var workbook = new Workbook(inputPath, loadOptions);

            // Iterate through each worksheet in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                var sheet = workbook.Worksheets[i];

                // Set image options: enable transparent background for PNG output
                var imgOptions = new ImageOrPrintOptions
                {
                    Transparent = true // Enables transparent background for PNG
                };

                try
                {
                    // Render the worksheet to an image using the specified options
                    var sheetRender = new SheetRender(sheet, imgOptions);
                    string outputFile = $"Sheet{i + 1}.png";

                    // Export the first page (index 0) of the sheet
                    sheetRender.ToImage(0, outputFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to render sheet {i + 1}: {ex.Message}");
                }
            }

            Console.WriteLine("PNG images with transparent background have been generated.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
