// Title: Create a Line Chart from a Smart‑Marker Template with WorkbookDesigner (Aspose.Cells C#)
// Description: Load an Excel template that contains smart markers, bind a collection of objects to the marker name, process the markers with WorkbookDesigner, compute the populated range, add a Line chart, set titles, and save the workbook.
// Keywords: Aspose.Cells | C# | WorkbookDesigner | Smart Markers | line chart | Excel template | chart automation | dynamic data range | example code | GitHub
// Common Searches: Aspose.Cells line chart from smart marker template | WorkbookDesigner add chart after processing markers | C# generate line chart using smart markers | Aspose.Cells dynamic chart range example | How to bind data to smart markers in Aspose.Cells
// Developer Intent: Generate a line chart by processing smart markers with WorkbookDesigner in Aspose.Cells.
// Use Cases: Automatically create monthly sales line charts from reusable Excel templates. | Produce charts for varying data sets without manually adjusting cell references. | Integrate chart generation into scheduled report‑building pipelines.
// AI Prompts: Write C# code that loads an Excel file with smart markers, binds a list of objects, processes the markers using WorkbookDesigner, and inserts a Line chart based on the populated data. | Explain how to determine the data range after WorkbookDesigner.Process() and configure NSeries and CategoryData for a line chart in Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSmartMarkerLineChart
{
    // Simple data class representing a data point for the chart
    // Load an Excel template that contains smart markers, bind a collection of objects to the marker name, process the markers with WorkbookDesigner, compute the populated range, add a Line chart, set titles, and save the workbook.
    public class DataPoint
    {
        public string Category { get; set; }
        public double Value { get; set; }

        public DataPoint(string category, double value)
        {
            Category = category;
            Value = value;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the workbook template that contains smart markers.
            // The template should have smart markers like "&Data.Category" and "&Data.Value"
            // placed in the first row (e.g., A1 and B1) and a named range "_CellsSmartMarkers".
            string templatePath = "Template.xlsx";

            // Load the template workbook.
            Workbook workbook = new Workbook(templatePath);

            // Prepare sample data that will replace the smart markers.
            List<DataPoint> data = new List<DataPoint>
            {
                new DataPoint("Jan", 120),
                new DataPoint("Feb", 150),
                new DataPoint("Mar", 170),
                new DataPoint("Apr", 130),
                new DataPoint("May", 190),
                new DataPoint("Jun", 210),
                new DataPoint("Jul", 180),
                new DataPoint("Aug", 160),
                new DataPoint("Sep", 200),
                new DataPoint("Oct", 220),
                new DataPoint("Nov", 190),
                new DataPoint("Dec", 230)
            };

            // Initialize WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook)
            {
                // Process the whole sheet at once (LineByLine = false) because the template uses a range.
                LineByLine = false
            };

            // Bind the data source to the smart marker name used in the template.
            designer.SetDataSource("Data", data);

            // Process the smart markers. The data will be populated into the worksheet.
            designer.Process();

            // After processing, the data will be placed starting from row 2 (A2, B2, ...).
            // Add a line chart that uses this populated data.
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row of data (header row + data count).
            int lastDataRow = 1 + data.Count; // Row index is zero‑based; row 0 is header.

            // Add a line chart to the worksheet.
            // Parameters: chart type, top row, left column, bottom row, right column.
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart.
            // Category (X‑axis) data: A2:A{lastDataRow}
            // Series (Y‑axis) data: B2:B{lastDataRow}
            string categoryRange = $"A2:A{lastDataRow}";
            string valuesRange = $"B2:B{lastDataRow}";
            chart.NSeries.Add(valuesRange, true);
            chart.NSeries.CategoryData = categoryRange;

            // Optional: set chart title and axis titles.
            chart.Title.Text = "Monthly Sales";
            chart.CategoryAxis.Title.Text = "Month";
            chart.ValueAxis.Title.Text = "Sales";

            // Save the resulting workbook.
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook with line chart saved to '{outputPath}'.");
        }
    }
}
