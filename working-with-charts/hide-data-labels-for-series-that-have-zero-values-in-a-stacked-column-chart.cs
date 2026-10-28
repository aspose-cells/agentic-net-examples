// Title: How to hide data labels for zero‑value points in a stacked column chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide data label for zero value chart point | remove zero value labels from stacked column chart Aspose.Cells | conditional visibility of chart data labels in Aspose.Cells .NET | C# Aspose.Cells hide labels on zero‑valued series points | Excel stacked column chart suppress zero labels with Aspose.Cells
// Tags: Aspose.Cells stacked column chart zero label suppression | C# chart data label conditional visibility | Aspose.Cells data labels hide zero values | Excel chart point label control Aspose.Cells | Aspose.Cells series point label customization

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data for two product series, inserts a stacked column chart, enables data labels for all series, then iterates each series point to hide the label when the point value is zero (ensuring a DataLabel object exists before toggling visibility), and finally saves the file as an XLSX workbook.
class HideZeroValueDataLabels
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Series 1
            sheet.Cells["B1"].PutValue("Product A");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(0);   // Zero value – label should be hidden
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["B5"].PutValue(130);

            // Series 2
            sheet.Cells["C1"].PutValue("Product B");
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["C3"].PutValue(95);
            sheet.Cells["C4"].PutValue(0);   // Zero value – label should be hidden
            sheet.Cells["C5"].PutValue(110);

            // Add a stacked column chart (using Column chart type for compatibility)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the series to the chart (true = use categories from column A)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.Add("C2:C5", true);

            // Enable data labels for all series
            foreach (var seriesObj in chart.NSeries)
            {
                dynamic series = seriesObj; // Use dynamic to avoid compile‑time type issues
                series.DataLabels.IsValueShown = true; // Show values by default
            }

            // Hide data labels where the point value is zero
            foreach (var seriesObj in chart.NSeries)
            {
                dynamic series = seriesObj;
                for (int i = 0; i < series.Values.Count; i++)
                {
                    object valObj = series.Values[i];
                    double val = 0;
                    if (valObj != null && double.TryParse(valObj.ToString(), out double parsed))
                    {
                        val = parsed;
                    }

                    // If the value is zero, hide its label
                    if (Math.Abs(val) < double.Epsilon)
                    {
                        // Ensure a DataLabel exists for the point
                        if (i >= series.DataLabels.Count)
                        {
                            series.DataLabels.Add();
                        }
                        series.DataLabels[i].IsValueShown = false;
                    }
                }
            }

            // Save the workbook
            string outputPath = "StackedColumn_HideZeroLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
