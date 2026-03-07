using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ShowHideGridlinesDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data to visualize gridlines
            worksheet.Cells["A1"].PutValue("Gridlines Demo");
            worksheet.Cells["A2"].PutValue("Row 2");
            worksheet.Cells["B2"].PutValue(123);
            worksheet.Cells["A3"].PutValue("Row 3");
            worksheet.Cells["B3"].PutValue(456);
            worksheet.AutoFitColumns();

            // Show gridlines
            worksheet.IsGridlinesVisible = true;
            workbook.Save("GridlinesVisible.xlsx");

            // Hide gridlines
            worksheet.IsGridlinesVisible = false;
            workbook.Save("GridlinesHidden.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShowHideGridlinesDemo.Run();
        }
    }
}