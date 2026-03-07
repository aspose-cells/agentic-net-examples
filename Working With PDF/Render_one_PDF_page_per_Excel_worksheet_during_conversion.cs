using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class OnePdfPagePerSheetDemo
    {
        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate first worksheet with sample data
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            for (int i = 0; i < 20; i++)
            {
                sheet1.Cells[i, 0].PutValue($"Row {i + 1} in Sheet1");
            }

            // Add a second worksheet and populate it
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            for (int i = 0; i < 30; i++)
            {
                sheet2.Cells[i, 0].PutValue($"Row {i + 1} in Sheet2");
            }

            // Set PDF save options to render each worksheet on a separate page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // One PDF page per worksheet
            };

            // Save the workbook as PDF; each worksheet becomes its own page
            workbook.Save("output.pdf", pdfOptions);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            OnePdfPagePerSheetDemo.Run();
        }
    }
}