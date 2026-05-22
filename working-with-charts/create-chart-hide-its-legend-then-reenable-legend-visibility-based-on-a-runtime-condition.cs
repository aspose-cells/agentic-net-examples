using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main(string[] args)
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the legend initially
        chart.ShowLegend = false;
        Console.WriteLine("Legend initially hidden: " + chart.ShowLegend);

        // Runtime condition to decide whether to show the legend again
        // Example: if any command‑line arguments are supplied, show the legend
        bool showLegendCondition = args.Length > 0;

        if (showLegendCondition)
        {
            chart.ShowLegend = true;
            Console.WriteLine("Condition met – legend re‑enabled: " + chart.ShowLegend);
        }

        // Save the workbook to a file
        workbook.Save("ChartLegendConditional.xlsx");
    }
}