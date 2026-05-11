using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    public class ExportCombinedWorkbookToPdf
    {
        public static void Run()
        {
            Workbook sourceWorkbook = new Workbook("SourceWorkbook.xlsx");
            Workbook secondWorkbook = new Workbook("SecondWorkbook.xlsx");
            sourceWorkbook.Combine(secondWorkbook);
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                CheckWorkbookDefaultFont = true,
                ExportDocumentStructure = true
            };
            sourceWorkbook.Save("CombinedWorkbook.pdf", pdfOptions);
            Console.WriteLine("Combined workbook exported to PDF successfully.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportCombinedWorkbookToPdf.Run();
        }
    }
}