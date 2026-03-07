using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSlicerToPdf
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source Excel file (XLSX) that contains the slicer
            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SampleWithSlicer.xlsx");

            // Desired output PDF file path
            string destPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SlicerOutput.pdf");

            // If the source file does not exist, create a simple workbook as a fallback
            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue("Sample Data");
                ws.Cells["A2"].PutValue(123);
                ws.Cells["A3"].PutValue(456);
                tempWb.Save(sourcePath, SaveFormat.Xlsx);
            }

            // Load the workbook
            var workbook = new Workbook(sourcePath);

            // Save the workbook as PDF (slicers are rendered automatically if present)
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed. PDF saved to: " + destPath);
        }
    }
}