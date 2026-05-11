using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            AutoFitAllRowsInMergedWorkbook.Run();
        }
    }

    public class AutoFitAllRowsInMergedWorkbook
    {
        public static void Run()
        {
            string inputPath = "MergedWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
            };

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.AutoFitRows(options);
            }

            string outputPath = "MergedWorkbook_AutoFitted.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}