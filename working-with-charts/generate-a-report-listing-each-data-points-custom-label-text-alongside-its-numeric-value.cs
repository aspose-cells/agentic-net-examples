using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class CustomDataLabelReport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate sample data for the chart
            // -------------------------------------------------
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("Item 1");
            dataSheet.Cells["A3"].PutValue("Item 2");
            dataSheet.Cells["A4"].PutValue("Item 3");
            dataSheet.Cells["B2"].PutValue(150);
            dataSheet.Cells["B3"].PutValue(250);
            dataSheet.Cells["B4"].PutValue(350);

            // -------------------------------------------------
            // Add a column chart and bind the data range
            // -------------------------------------------------
            int chartIndex = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = dataSheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // -------------------------------------------------
            // Enable data labels for the first series
            // -------------------------------------------------
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;        // Show numeric values

            // -------------------------------------------------
            // Set a custom label text for each data point
            // -------------------------------------------------
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];
                point.DataLabels.IsAutoText = false;               // Disable auto‑generated text
                point.DataLabels.Text = $"Label_{i + 1}";          // Custom label
            }

            // -------------------------------------------------
            // Create a separate worksheet to act as the report
            // -------------------------------------------------
            Worksheet reportSheet = workbook.Worksheets.Add("Report");
            reportSheet.Cells["A1"].PutValue("Custom Label");
            reportSheet.Cells["B1"].PutValue("Numeric Value");

            // -------------------------------------------------
            // Fill the report with each point's custom label and its value
            // -------------------------------------------------
            for (int i = 0; i < series.Points.Count; i++)
            {
                ChartPoint point = series.Points[i];
                string customLabel = point.DataLabels.Text;               // Retrieve custom label
                double numericValue = Convert.ToDouble(point.YValue);     // Retrieve numeric value safely

                int row = i + 2; // Data starts from row 2
                reportSheet.Cells[row, 0].PutValue(customLabel);
                reportSheet.Cells[row, 1].PutValue(numericValue);
            }

            // -------------------------------------------------
            // Save the workbook containing the chart and the report
            // -------------------------------------------------
            string outputPath = "CustomDataLabelReport.xlsx";

            // Ensure we don't attempt to overwrite a locked file
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}