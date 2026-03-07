using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class PageOrientationDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set page orientation to Landscape
            worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Add sample data
            worksheet.Cells["A1"].PutValue("Landscape Orientation Demo");
            for (int i = 2; i <= 10; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
            }

            // Save with Landscape orientation
            workbook.Save("PageOrientation_Landscape.xlsx", SaveFormat.Xlsx);

            // Change orientation to Portrait
            worksheet.PageSetup.Orientation = PageOrientationType.Portrait;

            // Save with Portrait orientation
            workbook.Save("PageOrientation_Portrait.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main()
        {
            PageOrientationDemo.Run();
        }
    }
}