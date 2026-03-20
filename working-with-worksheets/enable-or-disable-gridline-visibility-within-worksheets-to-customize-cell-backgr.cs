using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class GridlinesVisibilityDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Enable gridlines visibility
            worksheet.IsGridlinesVisible = true;
            // Add sample data to see the gridlines
            worksheet.Cells["A1"].PutValue("Gridlines Enabled");
            worksheet.Cells["A2"].PutValue("Visible in the sheet");

            // Save the workbook with gridlines visible
            workbook.Save("GridlinesEnabled.xlsx");

            // Disable gridlines visibility
            worksheet.IsGridlinesVisible = false;
            // Add more data to illustrate the change
            worksheet.Cells["A4"].PutValue("Gridlines Disabled");
            worksheet.Cells["A5"].PutValue("Not visible in the sheet");

            // Save the workbook with gridlines hidden
            workbook.Save("GridlinesDisabled.xlsx");

            Console.WriteLine("Workbooks saved with gridlines enabled and disabled.");
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