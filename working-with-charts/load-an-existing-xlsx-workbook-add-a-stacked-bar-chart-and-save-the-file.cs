using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddStackedBarChart
{
    static void Main()
    {
        // Load an existing XLSX workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Get the first worksheet (you can change the index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // OPTIONAL: add sample data if the workbook does not already contain it
        // sheet.Cells["A1"].PutValue("Category");
        // sheet.Cells["B1"].PutValue("Value");
        // for (int i = 2; i <= 6; i++)
        // {
        //     sheet.Cells[$"A{i}"].PutValue("Item " + (i - 1));
        //     sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        // }

        // Add a stacked bar chart to the worksheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 1, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (adjust the range to match your data)
        chart.NSeries.Add("=Sheet1!$A$1:$B$5", true);

        // Save the modified workbook
        string outputFile = "output.xlsx";
        workbook.Save(outputFile);
    }
}