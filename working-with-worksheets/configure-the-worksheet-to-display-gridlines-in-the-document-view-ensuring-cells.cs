using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class GridlinesDemo
    {
        public static void Run()
        {
            // Create a new workbook (create lifecycle)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Enable gridlines in the document view
            worksheet.IsGridlinesVisible = true;

            // Add sample data so the gridlines are visible
            worksheet.Cells["A1"].PutValue("Gridlines Demo");
            worksheet.Cells["A2"].PutValue("Sample");
            worksheet.Cells["B1"].PutValue("Demo");
            worksheet.Cells["B2"].PutValue("Data");

            // Adjust column widths for better visibility
            worksheet.AutoFitColumns();

            // Save the workbook (save lifecycle)
            workbook.Save("GridlinesDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GridlinesDemo.Run();
        }
    }
}