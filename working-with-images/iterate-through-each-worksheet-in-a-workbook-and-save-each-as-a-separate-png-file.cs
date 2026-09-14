// Title: Save each worksheet of an Excel workbook as a separate PNG image using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, iterates over all worksheets, and writes each sheet to an individual PNG file using SheetRender. | Show how to configure ImageOrPrintOptions for PNG output in single‑page‑per‑sheet mode and add checks for a missing input file plus per‑sheet error handling. | Provide a sample that logs the worksheet name and the path of the created PNG while gracefully handling rendering exceptions.
// Common Searches: Aspose.Cells C# export each sheet to separate PNG files | How to loop through worksheets and save as PNG using Aspose.Cells | C# render Excel worksheets to images one page per sheet | ImageOrPrintOptions one page per sheet Aspose.Cells example | Save Excel workbook sheets as individual PNG images in .NET
// Tags: Aspose.Cells worksheet to PNG conversion | C# iterate workbook worksheets export images | ImageOrPrintOptions PNG single-page mode | SheetRender render sheet as PNG | handle missing Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // This C# program checks for the presence of an input .xlsx file, loads it with Aspose.Cells, sets ImageOrPrintOptions to render each sheet as a single PNG page, iterates through every worksheet, uses SheetRender to create a separate PNG per sheet, and logs successes or any rendering errors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Configure image rendering options (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true // Render each sheet as a single page
            };

            // Iterate through each worksheet and save it as a separate PNG file
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                string sheetName = sheet.Name;

                try
                {
                    // Render the worksheet to an image
                    SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                    string outputPath = $"{sheetName}.png";
                    sheetRender.ToImage(0, outputPath);
                    Console.WriteLine($"Saved sheet '{sheetName}' as '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to render sheet '{sheetName}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
