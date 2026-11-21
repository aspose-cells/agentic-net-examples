// Title: Generate an Excel worksheet that lists custom label text with corresponding numeric values from a chart using Aspose.Cells for .NET
// Common Searches: how to add a separate report sheet with custom labels and values from a chart in Aspose.Cells C# | Aspose.Cells C# create worksheet that lists predefined labels and chart data values | export chart series numeric values to a new worksheet using Aspose.Cells for .NET | C# Aspose.Cells generate Excel file with column chart and a summary table of labels and values | save workbook with a custom label report sheet Aspose.Cells .NET example
// Tags: generate report worksheet Aspose.Cells | custom label and value table Excel C# | add column chart with data labels Aspose.Cells | export chart series to separate sheet Aspose.Cells | save workbook as .xlsx Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomLabelReport
{
    // The example creates a new workbook, inserts sample category and numeric data, builds a column chart displaying the values, then adds a 'Report' worksheet that lists predefined custom labels alongside their corresponding numeric values, and finally saves the workbook as CustomLabelReport.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // 1. Prepare sample data in the first worksheet
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Header
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Value");

                // Sample data rows
                string[] categories = { "A", "B", "C", "D" };
                double[] values = { 120.5, 85.3, 150.0, 95.7 };

                for (int i = 0; i < categories.Length; i++)
                {
                    dataSheet.Cells[i + 2, 0].PutValue(categories[i]); // Column A
                    dataSheet.Cells[i + 2, 1].PutValue(values[i]);    // Column B
                }

                // -------------------------------------------------
                // 2. Add a column chart based on the sample data
                // -------------------------------------------------
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = dataSheet.Charts[chartIndex];
                chart.Title.Text = "Sample Column Chart";

                // Set the data range for the series (values)
                chart.NSeries.Add("B2:B5", true);
                // Set the category (X) axis labels
                chart.NSeries.CategoryData = "A2:A5";

                // -------------------------------------------------
                // 3. Configure data labels for the series
                // -------------------------------------------------
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;          // Show numeric values
                series.DataLabels.ShowCategoryName = false; // Hide category names

                // Note: CustomLabels collection is not available in the current API version,
                // and PositionType enum may not be supported; therefore these settings are omitted.

                // -------------------------------------------------
                // 4. Generate a report worksheet listing custom label and numeric value
                // -------------------------------------------------
                Worksheet reportSheet = workbook.Worksheets.Add("Report");
                reportSheet.Cells["A1"].PutValue("Custom Label");
                reportSheet.Cells["B1"].PutValue("Numeric Value");

                string[] customLabels = { "Alpha", "Beta", "Gamma", "Delta" };
                for (int i = 0; i < customLabels.Length; i++)
                {
                    reportSheet.Cells[i + 2, 0].PutValue(customLabels[i]); // Column A
                    reportSheet.Cells[i + 2, 1].PutValue(values[i]);      // Column B
                }

                // -------------------------------------------------
                // 5. Save the workbook to a file
                // -------------------------------------------------
                string outputPath = "CustomLabelReport.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
