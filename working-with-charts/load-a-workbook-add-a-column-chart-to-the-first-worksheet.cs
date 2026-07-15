using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddColumnChartExample
{
    static void Main()
    {
        // Load an existing workbook from file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data if the sheet is empty (optional)
        if (sheet.Cells["A1"].Value == null)
        {
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 1; i <= 5; i++)
            {
                sheet.Cells[$"A{i + 1}"].PutValue($"Item {i}");
                sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
            }
        }

        // Add a column chart to the worksheet (rows 5‑20, columns 0‑5)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (including headers)
        chart.SetChartDataRange("A1:B6", true);

        // Set a title for the chart
        chart.Title.Text = "Sample Column Chart";

        // Save the workbook with the new chart
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}