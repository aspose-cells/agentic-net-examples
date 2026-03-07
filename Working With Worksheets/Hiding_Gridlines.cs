using System;
using Aspose.Cells;

namespace HideGridlinesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide gridlines in the worksheet
            worksheet.IsGridlinesVisible = false;

            // Add sample data to illustrate the result
            worksheet.Cells["A1"].PutValue("Gridlines are hidden");
            worksheet.Cells["A2"].PutValue("No gridlines will be shown");

            // Adjust column widths for better visibility
            worksheet.AutoFitColumns();

            // Save the workbook to a file
            workbook.Save("HideGridlinesDemo.xlsx");
        }
    }
}