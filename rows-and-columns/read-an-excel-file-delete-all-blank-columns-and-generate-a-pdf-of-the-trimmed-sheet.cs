using System;
using Aspose.Cells;

namespace AsposeCellsTrimAndExportPdf
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourceFile = "input.xlsx";

            // Path for the resulting PDF file
            string pdfFile = "trimmed_output.pdf";

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(sourceFile);

            // Access the first worksheet (you can change the index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Delete all blank columns in the worksheet
            worksheet.Cells.DeleteBlankColumns();

            // Save the trimmed workbook as a PDF document
            workbook.Save(pdfFile, SaveFormat.Pdf);

            Console.WriteLine("Blank columns removed and PDF generated successfully.");
        }
    }
}