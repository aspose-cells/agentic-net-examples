using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsAddInRender
{
    class Program
    {
        static void Main()
        {
            // Paths for the source XLSB workbook and the resulting PDF file
            string sourceXlsb = "input.xlsb";
            string outputPdf = "output.pdf";

            // Load the XLSB workbook
            Workbook workbook = new Workbook(sourceXlsb);

            // Apply a custom scaling factor (80%) to each worksheet.
            // The Zoom property expects a percentage value (e.g., 80 for 80%).
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.Zoom = 80; // 80% scaling
            }

            // Save the modified workbook back to XLSB using XlsbSaveOptions.
            // This ensures the scaling setting is persisted before conversion.
            XlsbSaveOptions xlsbOptions = new XlsbSaveOptions();
            workbook.Save(sourceXlsb, xlsbOptions);

            // Convert the updated XLSB workbook to PDF.
            // ConversionUtility handles the format conversion.
            ConversionUtility.Convert(sourceXlsb, outputPdf);

            Console.WriteLine("Conversion completed with custom scaling (0.8).");
        }
    }
}