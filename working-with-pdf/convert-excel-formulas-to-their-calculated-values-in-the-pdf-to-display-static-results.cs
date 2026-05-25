using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertFormulasToStaticPdf
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add a formula that sums the three values
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Configure PDF save options to calculate formulas before saving
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // When true, formulas are evaluated and their results are written to the PDF
                CalculateFormula = true
            };

            // Save the workbook as a PDF with calculated values (static results)
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, pdfOptions);

                // Write the PDF to a file for verification
                string outputPath = "FormulasCalculated.pdf";
                File.WriteAllBytes(outputPath, pdfStream.ToArray());
                Console.WriteLine($"PDF generated at '{Path.GetFullPath(outputPath)}' with static formula results.");
            }
        }
    }
}