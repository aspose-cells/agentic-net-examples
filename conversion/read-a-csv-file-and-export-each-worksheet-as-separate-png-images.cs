// Title: Convert a CSV file to individual PNG images for each worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a CSV file into an Aspose.Cells workbook and saves every worksheet as an individual PNG file. | Demonstrate how to set up image rendering options and generate PNG images for each sheet in a .NET project using Aspose.Cells. | Provide a script that creates an output folder, imports CSV data, iterates through all worksheets, and produces separate PNG screenshots with Aspose.Cells.
// Common Searches: Aspose.Cells C# export each worksheet to PNG image | How to render CSV data as PNG files using Aspose.Cells .NET | C# convert CSV to Excel then save sheets as PNG with Aspose | OnePagePerSheet option Aspose.Cells render multiple sheets to separate images | Generate PNG screenshots of Excel worksheets from CSV in C#
// Tags: import CSV into Aspose.Cells workbook C# | SheetRender export worksheet to PNG | ImageOrPrintOptions OnePagePerSheet usage | save each Excel sheet to separate PNG files | convert CSV to XLSX then render PNG with Aspose.Cells | batch render multiple worksheets to PNG in .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsCsvToPng
{
    // // This C# example reads a CSV file into a new Aspose.Cells workbook, creates an output directory, configures ImageOrPrintOptions for PNG with OnePagePerSheet, iterates through all worksheets, and uses SheetRender to generate a separate PNG image for each sheet, finally saving the workbook as an XLSX file for reference.
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvFilePath = "input.csv";

            // Directory where PNG images will be saved
            string outputDir = "output_images";
            Directory.CreateDirectory(outputDir);

            // Create a new workbook (empty)
            Workbook workbook = new Workbook();

            // Import the CSV data into the first worksheet (A1 cell)
            // Using comma as delimiter, convert numeric data, start at row 0, column 0
            workbook.Worksheets[0].Cells.ImportCSV(csvFilePath, ",", true, 0, 0);

            // If the CSV should be split into multiple worksheets, add that logic here.
            // For this example we assume a single worksheet.

            // Configure image rendering options for PNG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true // render the whole sheet as one page
            };

            // Iterate through each worksheet and render it to a PNG file
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Create a SheetRender for the current worksheet
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                // Since OnePagePerSheet = true, there will be only one page (index 0)
                string imagePath = Path.Combine(outputDir, $"Sheet{sheetIndex + 1}.png");

                // Render the page to a PNG file
                sheetRender.ToImage(0, imagePath);

                // Release resources used by SheetRender
                sheetRender.Dispose();

                Console.WriteLine($"Worksheet '{sheet.Name}' rendered to: {imagePath}");
            }

            // Optionally, save the workbook for reference
            workbook.Save("ConvertedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
