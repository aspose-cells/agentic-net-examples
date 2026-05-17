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

        // Simulate data that would be imported from an XML document.
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Item1");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Item2");
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["A4"].PutValue("Item3");
        sheet.Cells["B4"].PutValue(150);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 2, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart, plotting by column
        chart.SetChartDataRange("A1:B4", true);

        // Customize chart appearance
        chart.Title.Text = "Totals from XML";
        chart.Style = 2; // Built‑in style

        // Calculate the chart layout before saving
        chart.Calculate();

        // Save the workbook with the embedded chart
        workbook.Save("XmlMappedChart.xlsx", SaveFormat.Xlsx);
    }
}