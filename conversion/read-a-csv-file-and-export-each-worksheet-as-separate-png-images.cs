// Title: C# – Convert CSV to XLSX and Export Each Worksheet as PNG with Aspose.Cells
// Description: The sample converts a CSV file to a temporary XLSX workbook, loads it with Aspose.Cells, and iterates through every worksheet. For each sheet it configures ImageOrPrintOptions for PNG, renders the first page via SheetRender, saves the image as "Sheet_{n}.png", and finally deletes the temporary file.
// Keywords: Aspose.Cells | C# CSV to PNG | convert CSV to XLSX | SheetRender PNG export | ImageOrPrintOptions | export worksheet as image | batch worksheet image generation | temporary workbook cleanup | .NET spreadsheet rendering
// Common Searches: Aspose.Cells convert CSV to PNG per sheet | C# export each worksheet as separate PNG | how to render Excel sheet to image using Aspose.Cells | remove temporary XLSX after conversion Aspose | batch generate PNGs from CSV data C#
// Developer Intent: Create a PNG file for every worksheet produced from a CSV source.
// Use Cases: Generate visual previews of CSV‑derived sheets for dashboards or reports. | Automate image asset creation for web pages that display spreadsheet data. | Build a nightly batch that turns CSV datasets into individual PNG files for documentation.
// AI Prompts: Write C# code that reads a CSV, converts it to a workbook, and saves each worksheet as a PNG using Aspose.Cells. | Explain the role of ImageOrPrintOptions and SheetRender when rendering a worksheet to a PNG image. | Suggest how to name PNG files with the worksheet title and organize them into separate folders.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Utility;

namespace AsposeCellsCsvToPng
{
    // The sample converts a CSV file to a temporary XLSX workbook, loads it with Aspose.Cells, and iterates through every worksheet. For each sheet it configures ImageOrPrintOptions for PNG, renders the first page via SheetRender, saves the image as "Sheet_{n}.png", and finally deletes the temporary file.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Temporary XLSX file that will hold the converted CSV data
            string tempXlsxPath = "temp_converted.xlsx";

            // Convert CSV to XLSX using the provided ConversionUtility rule
            ConversionUtility.Convert(csvPath, tempXlsxPath);

            // Load the workbook from the converted XLSX file
            Workbook workbook = new Workbook(tempXlsxPath);

            // Iterate through each worksheet and export it as a separate PNG image
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                // Configure image rendering options (PNG format, one page per sheet)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true
                };

                // Create a SheetRender instance for the current worksheet
                SheetRender sheetRender = new SheetRender(workbook.Worksheets[sheetIndex], imgOptions);

                // Render the first (and only) page of the sheet to a PNG file
                string outputImagePath = $"Sheet_{sheetIndex + 1}.png";
                sheetRender.ToImage(0, outputImagePath);

                // Release resources used by SheetRender
                sheetRender.Dispose();

                Console.WriteLine($"Worksheet {sheetIndex + 1} exported to {outputImagePath}");
            }

            // Clean up the temporary XLSX file
            if (File.Exists(tempXlsxPath))
            {
                File.Delete(tempXlsxPath);
            }
        }
    }
}
