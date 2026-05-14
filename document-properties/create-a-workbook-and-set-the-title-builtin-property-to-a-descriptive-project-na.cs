using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetWorkbookTitleDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();
            workbook.BuiltInDocumentProperties.Title = "Project XYZ – Quarterly Report";
            workbook.Save("ProjectXYZ_QuarterlyReport.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetWorkbookTitleDemo.Run();
        }
    }
}