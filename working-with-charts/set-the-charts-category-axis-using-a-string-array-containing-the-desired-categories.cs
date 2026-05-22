using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCategoryAxisDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample numeric data for the chart (values in column B)
            sheet.Cells["A1"].PutValue("Category");   // placeholder for categories (not used)
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"B{i}"].PutValue(i * 10); // values 20,30,40,50
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values only)
            chart.NSeries.Add("B2:B5", true);

            // Set the category axis using a string array.
            // The string must be in the format "{'Cat1','Cat2',...}"
            chart.NSeries.CategoryData = "{'Jan','Feb','Mar','Apr'}";

            // Optional: give the series a name
            chart.NSeries[0].Name = "Sample Series";

            // Save the workbook
            workbook.Save("CategoryAxisWithStringArray.xlsx");
        }
    }
}