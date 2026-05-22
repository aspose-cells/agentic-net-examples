using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSmartMarkerLineChart
{
    // Sample data class representing a data point for the line chart
    public class ChartDataPoint
    {
        public string Category { get; set; }
        public double Value { get; set; }

        public ChartDataPoint(string category, double value)
        {
            Category = category;
            Value = value;
        }
    }

    public class GenerateLineChartWithSmartMarkers
    {
        public static void Run()
        {
            // Load the workbook template that contains smart markers.
            // The template should have a range where data will be populated, e.g. cells A2:B2 marked with "&Data.Category" and "&Data.Value".
            Workbook workbook = new Workbook("Template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare sample data for the smart markers.
            List<ChartDataPoint> data = new List<ChartDataPoint>
            {
                new ChartDataPoint("Jan", 120),
                new ChartDataPoint("Feb", 150),
                new ChartDataPoint("Mar", 170),
                new ChartDataPoint("Apr", 130),
                new ChartDataPoint("May", 190),
                new ChartDataPoint("Jun", 210)
            };

            // Bind the data source to the smart marker name used in the template (e.g., "Data").
            designer.SetDataSource("Data", data);

            // Process the smart markers – this will populate the worksheet with the data.
            designer.Process();

            // After processing, add a line chart that uses the populated data.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a line chart positioned from row 10, column 1 to row 30, column 8.
            int chartIndex = sheet.Charts.Add(ChartType.Line, 9, 0, 29, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart.
            // Assuming the template placed categories in column A and values in column B starting from row 2.
            // Adjust the range according to the actual number of data rows.
            string dataRange = $"=Sheet1!$A$2:$B${data.Count + 1}";
            chart.NSeries.Add(dataRange, true);

            // Optional: set chart title and enable legend.
            chart.Title.Text = "Monthly Sales";
            chart.ShowLegend = true;

            // Save the resulting workbook.
            workbook.Save("Result.xlsx");
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static void Main()
        {
            GenerateLineChartWithSmartMarkers.Run();
            Console.WriteLine("Line chart generated and saved to Result.xlsx");
        }
    }
}