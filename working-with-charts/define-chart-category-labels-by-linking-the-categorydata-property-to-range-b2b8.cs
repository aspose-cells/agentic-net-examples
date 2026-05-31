using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetCategoryDataDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: numeric values in column A and category labels in column B
        sheet.Cells["A1"].PutValue("Value");
        sheet.Cells["B1"].PutValue("Category");
        for (int i = 2; i <= 8; i++)
        {
            sheet.Cells[$"A{i}"].PutValue(i * 10);               // Example values: 20, 30, ...
            sheet.Cells[$"B{i}"].PutValue("Cat" + (i - 1));     // Category labels: Cat1, Cat2, ...
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the series data range (values) – vertical orientation
        chart.NSeries.Add("A2:A8", true);

        // Link the category axis labels to the range B2:B8
        chart.NSeries.CategoryData = "B2:B8";

        // Save the workbook to a file
        workbook.Save("CategoryDataLinked.xlsx");
    }
}