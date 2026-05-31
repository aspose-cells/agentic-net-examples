using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsShowZeroExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(0);
            sheet.Cells["A2"].PutValue(123);

            // Iterate through all worksheets and hide zero values
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.DisplayZeros = false; // Do not display zero values
            }

            // Save the workbook as PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}