using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class XmlMappedChartDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // Simulate data that would come from mapped XML nodes.
        // In a real scenario the XML map would populate these cells.
        // ------------------------------------------------------------
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Total");

        // Example categories and totals
        sheet.Cells["A2"].PutValue("North");
        sheet.Cells["B2"].PutValue(12500);
        sheet.Cells["A3"].PutValue("South");
        sheet.Cells["B3"].PutValue(9800);
        sheet.Cells["A4"].PutValue("East");
        sheet.Cells["B4"].PutValue(14300);
        sheet.Cells["A5"].PutValue("West");
        sheet.Cells["B5"].PutValue(11200);

        // Add a column chart to the worksheet (rows 7‑20, columns 2‑8)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 2, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (including headers)
        chart.SetChartDataRange("A1:B5", true);

        // Optional: set chart title and style
        chart.Title.Text = "Totals by Region";
        chart.Style = 2; // Built‑in style

        // Save the workbook with the embedded chart
        workbook.Save("XmlMappedChart.xlsx", SaveFormat.Xlsx);
    }
}