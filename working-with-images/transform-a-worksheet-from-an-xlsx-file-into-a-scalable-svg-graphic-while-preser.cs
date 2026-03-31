using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the generated SVG file
            string svgPath = "output.svg";

            // Load the workbook from the XLSX file (uses Aspose.Cells Workbook constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Get the first worksheet (or any specific worksheet you need)
            Worksheet worksheet = workbook.Worksheets[0];

            // Create SVG rendering options
            SvgImageOptions svgOptions = new SvgImageOptions();

            // Enable FitToViewPort to make the SVG scalable and fit the viewport
            svgOptions.FitToViewPort = true;

            // Initialize the sheet renderer with the worksheet and SVG options
            SheetRender sheetRender = new SheetRender(worksheet, svgOptions);

            // Render the first page of the worksheet to an SVG file
            // The ToImage overload saves directly to a file
            sheetRender.ToImage(0, svgPath);

            Console.WriteLine($"Worksheet has been successfully exported to SVG: {svgPath}");
        }
    }
}