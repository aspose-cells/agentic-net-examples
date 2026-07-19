// Title: Aspose.Cells .NET: Export Custom Chart Data Labels to a Summary Worksheet
// Description: C# code that creates a workbook, adds sample category/value data, builds a column chart, assigns a custom label to each chart point, and generates a separate "Report" sheet listing each custom label with its numeric value. The file is saved as CustomLabelReport.xlsx.
// Keywords: Aspose.Cells | C# | ASP.NET | chart data labels | custom labels | export chart points | Excel report generation | column chart | ChartPoint | Series data extraction
// Common Searches: Aspose.Cells retrieve custom label text from chart points | export chart point values to Excel using Aspose.Cells .NET | list chart data labels in a new worksheet C# | how to write chart point labels to cells with Aspose.Cells
// Developer Intent: Generate an Excel workbook that lists each chart point’s custom label together with its numeric value.
// Use Cases: Create a summary sheet that pairs custom‑labeled categories with their sales numbers for business reporting. | Export chart annotations to a tabular format for data validation or audit trails. | Produce a printable report that combines descriptive labels and measured values for presentation purposes.
// AI Prompts: Show how to include the series name in the report alongside each custom label and value. | Provide code to export the custom label report to a CSV file instead of an Excel worksheet. | Explain how to apply a specific number format (e.g., currency) to the numeric values in the report sheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomLabelReport
{
    // C# code that creates a workbook, adds sample category/value data, builds a column chart, assigns a custom label to each chart point, and generates a separate "Report" sheet listing each custom label with its numeric value. The file is saved as CustomLabelReport.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["A2"].PutValue("Alpha");
                dataSheet.Cells["A3"].PutValue("Beta");
                dataSheet.Cells["A4"].PutValue("Gamma");

                dataSheet.Cells["B1"].PutValue("Value");
                dataSheet.Cells["B2"].PutValue(150);
                dataSheet.Cells["B3"].PutValue(300);
                dataSheet.Cells["B4"].PutValue(450);

                // Add a column chart
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = dataSheet.Charts[chartIndex];

                // Set the data range for the series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Assign custom label text to each data point
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    // Disable auto-generated text
                    point.DataLabels.IsAutoText = false;
                    // Example custom label: combine category name with a prefix
                    string customLabel = $"Item {dataSheet.Cells[i + 2, 0].StringValue}";
                    point.DataLabels.Text = customLabel;
                }

                // Create a new worksheet to hold the report
                int reportSheetIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
                reportSheet.Name = "Report";

                // Write headers
                reportSheet.Cells["A1"].PutValue("Custom Label");
                reportSheet.Cells["B1"].PutValue("Numeric Value");

                // Populate the report with each point's custom label and its numeric value
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    // Custom label text
                    string label = point.DataLabels.Text;
                    // Numeric value of the point (cast to double)
                    double value = Convert.ToDouble(point.YValue);

                    // Write to the report sheet (starting from row 2)
                    int row = i + 2;
                    reportSheet.Cells[row, 0].PutValue(label);
                    reportSheet.Cells[row, 1].PutValue(value);
                }

                // Define output file path
                string outputPath = "CustomLabelReport.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
