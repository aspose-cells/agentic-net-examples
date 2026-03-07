using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfReview
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path where the resulting PDF will be saved
            string pdfPath = "output.pdf";

            // Load the workbook to evaluate page count before conversion
            Workbook workbook = new Workbook(sourcePath);

            // Create print options (default settings)
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions();

            // Get the total number of pages that will be rendered
            WorkbookPrintingPreview preview = new WorkbookPrintingPreview(workbook, printOptions);
            int pageCount = preview.EvaluatedPageCount;

            // Convert the Excel file to PDF using the provided ConversionUtility rule
            ConversionUtility.Convert(sourcePath, pdfPath);

            // Review information about the generated PDF
            Console.WriteLine($"Conversion completed successfully.");
            Console.WriteLine($"Source file: {sourcePath}");
            Console.WriteLine($"PDF file: {pdfPath}");
            Console.WriteLine($"Estimated page count: {pageCount}");
        }
    }
}