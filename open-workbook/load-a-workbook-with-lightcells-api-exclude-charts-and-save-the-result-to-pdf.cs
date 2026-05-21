using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsLightCellsExample
{
    class Program
    {
        static void Main()
        {
            // Input Excel file path (replace with your actual file)
            string inputPath = "input.xlsx";

            // Output PDF file path
            string outputPath = "output.pdf";

            // Load the workbook using LoadOptions (LightCells mode can be enabled via LoadOptions if needed)
            LoadOptions loadOptions = new LoadOptions(); // default options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Remove all charts from every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Remove charts one by one until none remain
                while (sheet.Charts.Count > 0)
                {
                    sheet.Charts.RemoveAt(0);
                }
            }

            // Configure PDF save options (optional settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: ignore rendering errors to ensure PDF is generated even if issues exist
                IgnoreError = true
            };

            // Save the modified workbook as PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF without charts at: {outputPath}");
        }
    }
}