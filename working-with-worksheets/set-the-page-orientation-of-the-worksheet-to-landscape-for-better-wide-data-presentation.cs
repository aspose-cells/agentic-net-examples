using System;
using Aspose.Cells;

namespace AsposeCellsOrientationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the page orientation to Landscape for better wide data presentation
            worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // (Optional) Add some sample data to visualize the effect
            worksheet.Cells["A1"].PutValue("Landscape Orientation Demo");
            for (int i = 1; i <= 10; i++)
            {
                worksheet.Cells[$"A{i + 1}"].PutValue($"Data Row {i}");
            }

            // Save the workbook to a file
            workbook.Save("LandscapeOrientationDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}