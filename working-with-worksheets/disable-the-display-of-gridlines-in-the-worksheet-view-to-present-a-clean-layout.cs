using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DisableGridlinesDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide gridlines in the worksheet view
            worksheet.IsGridlinesVisible = false;

            // Add sample data to verify the layout
            worksheet.Cells["A1"].PutValue("Gridlines are hidden");
            worksheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("GridlinesHidden.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DisableGridlinesDemo.Run();
        }
    }
}