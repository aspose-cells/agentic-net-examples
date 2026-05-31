using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data: categories in column A, values in column B
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series (values)
        chart.NSeries.Add("B2:B6", true);

        // Assign the cell range that will serve as the category axis values
        chart.NSeries.CategoryData = "A2:A6";

        // Save the workbook to a file
        workbook.Save("CategoryAxisDemo.xlsx");
    }
}