using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ShowRowColumnHeadersDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure row and column headers are visible in the worksheet view
            worksheet.IsRowColumnHeadersVisible = true;

            // Add some sample data
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue("Data1");
            worksheet.Cells["B2"].PutValue("Data2");

            // Save the workbook
            workbook.Save("ShowRowColumnHeadersDemo.xlsx");

            // Load the saved workbook to verify the property
            Workbook loadedWorkbook = new Workbook("ShowRowColumnHeadersDemo.xlsx");
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

            // Output the current state of the header visibility property
            Console.WriteLine("Row and Column Headers Visible: " + loadedWorksheet.IsRowColumnHeadersVisible);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShowRowColumnHeadersDemo.Run();
        }
    }
}