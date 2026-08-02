using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAxisCategoryDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: text categories in column A and numeric values in column B
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and the category axis (text labels)
            chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

            // Change the X axis (category axis) type to CategoryScale so that sequential text labels are shown
            chart.CategoryAxis.CategoryType = CategoryType.CategoryScale;

            // Optional: give the axis a title
            chart.CategoryAxis.Title.Text = "Categories";

            // Save the workbook to an XLSX file
            workbook.Save("ChartWithCategoryAxis.xlsx", SaveFormat.Xlsx);
        }
    }
}