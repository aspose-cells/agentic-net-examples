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

        // Populate sample data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(85);
        sheet.Cells["B4"].PutValue(65);

        // Add a pie chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the data labels of the first series
        DataLabels labels = chart.NSeries[0].DataLabels;

        // Show values and category names on the data labels
        labels.ShowValue = true;
        labels.ShowCategoryName = true;

        // Enable text wrapping for the data labels
        labels.IsTextWrapped = true;

        // Save the workbook to a file
        workbook.Save("PieChartDataLabelsWrapped.xlsx");
    }
}