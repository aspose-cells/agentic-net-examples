using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Utility;

namespace AsposeCellsAddInRenderDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Path where the resulting PDF will be saved
            string pdfPath = "output.pdf";

            // Load the workbook (create & load)
            Workbook workbook = new Workbook(sourcePath);

            // Prepare rendering options (default scaling, no explicit Zoom)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                // OnePagePerSheet = false (default) allows normal pagination
                // No Zoom is set, so default scaling will be used
            };

            // Create a SheetRender for the first worksheet to obtain the calculated page scale
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], renderOptions);

            // Retrieve the default page scale (calculated based on FitToPages settings)
            double defaultPageScale = sheetRender.PageScale;
            Console.WriteLine($"Default page scale (percentage): {defaultPageScale * 100}%");

            // Dispose the render object as it is no longer needed
            sheetRender.Dispose();

            // Convert the workbook to PDF using the ConversionUtility (save)
            ConversionUtility.Convert(sourcePath, pdfPath);

            Console.WriteLine($"Workbook successfully converted to PDF at: {pdfPath}");
        }
    }
}