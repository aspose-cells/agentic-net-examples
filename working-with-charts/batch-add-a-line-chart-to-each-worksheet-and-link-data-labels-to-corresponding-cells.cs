using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class BatchLineChartWithLinkedLabels
{
    static void Main()
    {
        // Create a new workbook and add a few worksheets for demonstration
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Process each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // -------------------------------------------------
            // 1. Prepare sample data (Category, Value, Label)
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Label");

            for (int row = 2; row <= 11; row++)
            {
                sheet.Cells[$"A{row}"].PutValue($"Item {row - 1}");
                sheet.Cells[$"B{row}"].PutValue(row * 10);                     // numeric value
                sheet.Cells[$"C{row}"].PutValue($"Lbl {row - 1}");            // label to link
            }

            // -------------------------------------------------
            // 2. Add a line chart to the worksheet
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define data range for the series and categories
            string sheetName = sheet.Name;
            chart.NSeries.Add($"={sheetName}!$B$2:$B$11", true);
            chart.NSeries.CategoryData = $"={sheetName}!$A$2:$A$11";

            // -------------------------------------------------
            // 3. Link data labels to the label cells (Column C)
            // -------------------------------------------------
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;                                 // show the numeric value
            series.DataLabels.LinkedSource = $"={sheetName}!$C$2:$C$11";        // link to label cells
            series.DataLabels.NumberFormatLinked = true;                       // keep number format in sync

            // Optional: position labels outside the line points
            series.DataLabels.Position = LabelPositionType.OutsideEnd;
        }

        // Save the workbook
        workbook.Save("BatchLineChartsWithLinkedLabels.xlsx", SaveFormat.Xlsx);
    }
}