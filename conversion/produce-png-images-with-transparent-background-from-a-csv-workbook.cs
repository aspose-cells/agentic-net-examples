using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsCsvToTransparentPng
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Path for the resulting PNG image
            string pngPath = "output.png";

            // Create an empty workbook
            Workbook workbook = new Workbook();

            // Import the CSV data into the first worksheet starting at cell A1 (row 0, column 0)
            // Using comma as delimiter and converting numeric strings to numbers
            workbook.Worksheets[0].Cells.ImportCSV(csvPath, ",", true, 0, 0);

            // Configure image rendering options
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,   // Output format PNG
                Transparent = true,          // Make background transparent
                OnePagePerSheet = true       // Render each sheet as a single page
            };

            // Create a workbook renderer with the workbook and rendering options
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);

            // Render the entire workbook to a PNG file with transparent background
            renderer.ToImage(pngPath);

            Console.WriteLine($"Workbook rendered to transparent PNG: {Path.GetFullPath(pngPath)}");
        }
    }
}