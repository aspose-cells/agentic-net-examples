using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class GridlinesVisibilityDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.IsGridlinesVisible = true;
            worksheet.Cells["A1"].PutValue("Gridlines Visible Example");
            worksheet.Cells["A2"].PutValue("This worksheet shows gridlines");
            worksheet.AutoFitColumns();
            workbook.Save("GridlinesVisibleDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GridlinesVisibilityDemo.Run();
        }
    }
}