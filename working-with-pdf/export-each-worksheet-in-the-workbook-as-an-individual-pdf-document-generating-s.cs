using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class ExportSheetsToSeparatePdf
    {
        public static void Run()
        {
            // Create a new workbook with sample data
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Populate each sheet with some data
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                sheet.Cells["A1"].PutValue($"Data from {sheet.Name}");
                sheet.Cells["A2"].PutValue(DateTime.Now);
            }

            // Loop through all worksheets and save each one as an individual PDF
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    SheetSet = new SheetSet(new int[] { i })
                };

                string outputFile = $"{workbook.Worksheets[i].Name}.pdf";
                workbook.Save(outputFile, pdfOptions);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportSheetsToSeparatePdf.Run();
        }
    }
}