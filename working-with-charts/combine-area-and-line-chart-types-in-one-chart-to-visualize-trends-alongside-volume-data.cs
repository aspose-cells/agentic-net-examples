using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ComboChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Month, Volume (area), Trend (line)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Volume");
            sheet.Cells["C1"].PutValue("Trend");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            double[] volume = { 120, 150, 180, 130, 170, 200 };
            double[] trend = { 130, 140, 160, 150, 165, 190 };

            for (int i = 0; i < months.Length; i++)
            {
                int row = i + 2; // data starts from row 2
                sheet.Cells[row, 0].PutValue(months[i]);   // Column A: Month
                sheet.Cells[row, 1].PutValue(volume[i]);   // Column B: Volume
                sheet.Cells[row, 2].PutValue(trend[i]);    // Column C: Trend
            }

            // Add a combo chart: start with Area type
            int chartIndex = sheet.Charts.Add(ChartType.Area, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Volume (Area) and Trend (Line)";

            // Set the category (X‑axis) data for the chart
            chart.NSeries.CategoryData = "A2:A7";

            // Add the first series (Volume) – default Area type
            chart.NSeries.Add("B2:B7", false);
            chart.NSeries[0].Name = "Volume";

            // Add the second series (Trend) and change its type to Line
            chart.NSeries.Add("C2:C7", false);
            chart.NSeries[1].Name = "Trend";
            chart.NSeries[1].Type = ChartType.Line;   // Convert this series to a line chart

            // Optional: display the line series on a secondary axis
            // Note: Some Aspose.Cells versions may not expose IsOnSecondaryAxis.
            // If needed, uncomment the line below when the property is available.
            // chart.NSeries[1].IsOnSecondaryAxis = true;

            // Determine output path and ensure directory exists
            string outputPath = "ComboChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the combo chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}