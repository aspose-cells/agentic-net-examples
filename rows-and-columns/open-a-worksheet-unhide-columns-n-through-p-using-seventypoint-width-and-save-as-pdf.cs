using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide columns N (index 13) through P (index 15) and set their width to 70 points
        worksheet.Cells.UnhideColumns(13, 3, 70);

        // Save the workbook as a PDF
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("output.pdf", pdfOptions);
    }
}