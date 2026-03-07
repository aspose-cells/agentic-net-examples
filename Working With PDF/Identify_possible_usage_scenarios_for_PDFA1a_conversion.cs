using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Utility;
using Aspose.Cells.Charts;

namespace AsposeCellsPdfA1aScenarios
{
    public class PdfA1aExamples
    {
        // Scenario 1: Directly save a workbook as PDF/A-1a
        public static void SaveWorkbookAsPdfA1a()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("PDF/A-1a compliance example");
            ws.Cells["A2"].PutValue(DateTime.Now);

            PdfSaveOptions options = new PdfSaveOptions();
            options.Compliance = PdfCompliance.PdfA1a;

            wb.Save("Workbook_PdfA1a.pdf", options);
        }

        // Scenario 2: Convert an existing Excel file to PDF/A-1a using ConversionUtility
        public static void ConvertExcelToPdfA1a()
        {
            string sourcePath = "Sample.xlsx";

            Workbook srcWb = new Workbook();
            srcWb.Worksheets[0].Cells["A1"].PutValue("Source Excel for conversion");
            srcWb.Save(sourcePath);

            LoadOptions loadOptions = new LoadOptions();
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.Compliance = PdfCompliance.PdfA1a;

            ConversionUtility.Convert(sourcePath, loadOptions, "Converted_PdfA1a.pdf", saveOptions);
        }

        // Scenario 3: Export a chart to PDF/A-1a by saving the whole workbook with compliance
        public static void ExportChartWithPdfA1a()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("Apples");
            ws.Cells["A3"].PutValue("Oranges");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["B2"].PutValue(50);
            ws.Cells["B3"].PutValue(30);

            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 10);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            PdfSaveOptions options = new PdfSaveOptions();
            options.Compliance = PdfCompliance.PdfA1a;

            wb.Save("Chart_PdfA1a.pdf", options);
        }

        public static void Main()
        {
            SaveWorkbookAsPdfA1a();
            ConvertExcelToPdfA1a();
            ExportChartWithPdfA1a();
            Console.WriteLine("All PDF/A-1a scenarios executed.");
        }
    }
}