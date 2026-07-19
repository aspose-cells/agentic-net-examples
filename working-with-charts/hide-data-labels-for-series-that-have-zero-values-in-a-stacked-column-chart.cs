// Title: C# Aspose.Cells: Hide Zero‑Value Data Labels in a Stacked Column Chart
// Description: Shows how to build a workbook, insert data that includes zeros, create a stacked column chart, turn on data labels, and then programmatically turn off the label for any point whose underlying cell value is zero by using the Series.Points.DataLabels.ShowValue property.
// Keywords: Aspose.Cells | C# | .NET | stacked column chart | hide zero data labels | conditional data label visibility | chart customization | Excel export | Series.Points.DataLabels | ShowValue property
// Common Searches: Aspose.Cells hide zero labels | C# stacked column chart hide data labels for zero values | remove zero value labels from Excel chart using Aspose | conditional data label visibility Aspose.Cells | how to suppress zero data labels in a chart
// Developer Intent: Programmatically suppress data labels for points with a zero value in a stacked column chart.
// Use Cases: Sales dashboards where categories with no sales should not display a label, keeping the chart clean. | Financial reports that hide zero‑balance entries in stacked column visualizations for better readability. | Automated Excel generation where only meaningful data points are labeled, reducing visual clutter.
// AI Prompts: Generate C# code with Aspose.Cells that hides data labels for zero‑value points in a stacked column chart. | Provide an Aspose.Cells example that iterates through series points and disables ShowValue when the cell value equals zero. | Explain how to apply conditional formatting to chart data labels based on cell values using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, insert data that includes zeros, create a stacked column chart, turn on data labels, and then programmatically turn off the label for any point whose underlying cell value is zero by using the Series.Points.DataLabels.ShowValue property.
    public class HideZeroValueDataLabelsInStackedColumnChart
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including zero values)
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            // Series 1 values
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(0);   // Zero value
            sheet.Cells["B4"].PutValue(20);

            // Series 2 values
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(0);   // Zero value

            // Add a stacked column chart
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (both series)
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for all series (initially show values)
            foreach (Series ser in chart.NSeries)
            {
                ser.DataLabels.ShowValue = true;
            }

            // Hide data labels for points with zero value
            int seriesIdx = 0; // 0 -> Series1 (column B), 1 -> Series2 (column C)
            foreach (Series ser in chart.NSeries)
            {
                int dataColumn = 1 + seriesIdx; // B=1, C=2 (zero‑based column index)
                for (int i = 0; i < ser.Points.Count; i++)
                {
                    // Row index for data starts at row 2 (index 1)
                    int dataRow = i + 1;
                    object cellObj = sheet.Cells[dataRow, dataColumn].Value;
                    double cellValue = 0;
                    if (cellObj != null && double.TryParse(cellObj.ToString(), out double parsed))
                    {
                        cellValue = parsed;
                    }

                    // If the cell value is zero, hide its data label
                    if (Math.Abs(cellValue) < double.Epsilon)
                    {
                        ser.Points[i].DataLabels.ShowValue = false;
                    }
                }
                seriesIdx++;
            }

            // Save the workbook
            string outputPath = "StackedColumn_HideZeroDataLabels.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}
