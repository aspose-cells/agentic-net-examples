using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfFlattenDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // using the default constructor

                // Add some sample data (optional, just for demonstration)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");
                sheet.Cells["B1"].PutValue(123);

                // Create PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // NOTE: In some versions of Aspose.Cells the FlattenAllAnnotations property is not available.
                // If your version supports it, you can uncomment the following line:
                // pdfOptions.FlattenAllAnnotations = true;

                // Optionally calculate formulas before saving
                workbook.CalculateFormula();

                // Save the workbook as a PDF using the options
                string outputPath = "FlattenedAnnotations.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook has been saved to PDF at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}