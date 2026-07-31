// Title: Add Data Labels to a Radar Chart from Transposed Row Data with Auto‑Fit Shapes – Aspose.Cells for .NET
// Description: This C# example demonstrates how to create an Excel workbook with Aspose.Cells, populate a table where categories are in column A and series values run across rows, insert a radar chart, set the data source by rows (transposed), enable data labels that show numeric values, automatically resize label shapes to fit the text, apply a RoundRect shape, activate radar axis labels, recalculate the chart layout, and save the file.
// Keywords: Aspose.Cells radar chart | C# radar chart data labels | set chart data range rows | auto fit data label shape | RoundRect data label Aspose.Cells | radar axis labels .NET | transposed data range Excel chart | Aspose.Cells .NET example | programmatic Excel radar chart | chart data labels Aspose.Cells
// Common Searches: Aspose.Cells show values on radar chart series | How to use rows as data source for radar chart in .NET | Resize data label shape to fit text Aspose.Cells | Enable category axis labels on radar chart Aspose.Cells | C# example for radar chart with auto‑fit labels
// Developer Intent: Create a radar chart that uses row‑based (transposed) data, displays value labels, and auto‑fits label shapes.
// Use Cases: Visualizing performance metrics across multiple categories with clear numeric labels. | Generating radar charts from tables where categories are listed vertically and series data is horizontal. | Producing professional‑looking Excel reports where data label shapes adapt to varying value lengths.
// AI Prompts: Write C# code using Aspose.Cells to add a radar chart, set its data range by rows, enable data labels with values, and auto‑fit the label shapes. | Show how to apply a RoundRect shape to data labels and turn on radar axis labels in an Aspose.Cells radar chart. | Explain the steps to transpose a cell range for a radar chart and customize data label appearance with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsRadarDemo
{
    // This C# example demonstrates how to create an Excel workbook with Aspose.Cells, populate a table where categories are in column A and series values run across rows, insert a radar chart, set the data source by rows (transposed), enable data labels that show numeric values, automatically resize label shapes to fit the text, apply a RoundRect shape, activate radar axis labels, recalculate the chart layout, and save the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (categories in column A, series values in rows B‑E)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Cat1");
                sheet.Cells["A3"].PutValue("Cat2");
                sheet.Cells["A4"].PutValue("Cat3");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["D1"].PutValue("Series3");
                sheet.Cells["E1"].PutValue("Series4");

                sheet.Cells["B2"].PutValue(4);
                sheet.Cells["C2"].PutValue(5);
                sheet.Cells["D2"].PutValue(3);
                sheet.Cells["E2"].PutValue(6);

                sheet.Cells["B3"].PutValue(2);
                sheet.Cells["C3"].PutValue(7);
                sheet.Cells["D3"].PutValue(4);
                sheet.Cells["E3"].PutValue(5);

                sheet.Cells["B4"].PutValue(5);
                sheet.Cells["C4"].PutValue(3);
                sheet.Cells["D4"].PutValue(6);
                sheet.Cells["E4"].PutValue(2);

                // Add a radar chart
                int chartIndex = sheet.Charts.Add(ChartType.Radar, 6, 0, 22, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Use rows as data source (set isVertical to false)
                chart.SetChartDataRange("A1:E4", false); // false = rows, true = columns

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;                     // Show the numeric values
                series.DataLabels.IsResizeShapeToFitText = true;        // Auto‑fit the label shape to its text
                series.DataLabels.ShapeType = DataLabelShapeType.RoundRect; // Example shape

                // Enable radar axis (category) labels
                series.HasRadarAxisLabels = true;

                // Recalculate chart layout before saving
                chart.Calculate();

                // Save the workbook
                string outputPath = "RadarChartWithDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
