// Title: Aspose.Cells C# – Add Data Labels to a Column Chart and Freeze Header Rows
// Description: Demonstrates how to create a workbook, populate a data sheet, add a column chart on a separate chart sheet, enable data labels (showing values, positioning them outside the columns, and applying custom font), freeze the first four rows of the source worksheet, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart data labels | column chart | freeze panes | freeze rows | chart sheet | Excel export | ShowValue | LabelPositionType | Aspose.Cells FreezePanes | chart series styling
// Common Searches: Aspose.Cells add data labels to chart | How to freeze rows in Aspose.Cells C# | Enable data labels on column chart Aspose.Cells | Create chart sheet with Aspose.Cells .NET | Freeze panes while scrolling in Excel using Aspose.Cells
// Developer Intent: Add data labels to a chart series and keep the label rows visible by freezing them in the worksheet.
// Use Cases: Show sales numbers directly on a column chart while keeping category headers in view during scrolling. | Generate a reporting workbook where chart labels stay static by freezing the top rows of the data sheet. | Create a chart sheet with styled data labels and lock header rows to act as a persistent legend for analysts.
// AI Prompts: Write C# code with Aspose.Cells that creates a line chart, enables data labels showing percentages, and freezes the first three rows of the data worksheet. | Provide an example of applying custom font styling to chart data labels and then freezing both rows and columns on the same sheet using Aspose.Cells. | Explain how to reference data from one worksheet in a chart placed on a separate chart sheet while keeping the source rows frozen with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsChartWithFrozenLabels
{
    // Demonstrates how to create a workbook, populate a data sheet, add a column chart on a separate chart sheet, enable data labels (showing values, positioning them outside the columns, and applying custom font), freeze the first four rows of the source worksheet, and save the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (data sheet)
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["A4"].PutValue("C");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["B4"].PutValue(30);

            // Add a chart sheet to hold the chart
            int chartSheetIndex = workbook.Worksheets.Add(SheetType.Chart);
            Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];

            // Create a column chart on the chart sheet
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = chartSheet.Charts[chartIndex];

            // Set the data source for the chart (referencing the data sheet)
            chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

            // Enable data labels for the first series and show the values
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.OutsideEnd; // optional positioning
            series.DataLabels.Font.Color = Color.Blue; // optional styling
            series.DataLabels.ApplyFont(); // apply font changes to all labels

            // Freeze the rows that contain the data labels (rows 1‑4) on the data sheet
            // This keeps the label rows visible while scrolling
            dataSheet.FreezePanes(5, 1, 4, 0); // Freeze first 4 rows, no columns frozen

            // Save the workbook
            workbook.Save("ChartWithDataLabelsAndFrozenRows.xlsx", SaveFormat.Xlsx);
        }
    }
}
