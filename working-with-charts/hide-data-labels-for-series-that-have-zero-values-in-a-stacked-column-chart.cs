// Title: Hide Zero-Value Data Labels in Stacked Column Chart – Aspose.Cells for .NET (C#)
// Description: This example builds a workbook, inserts sample data that includes zero values, creates a stacked column chart, enables data labels for all series, then scans each point, reads its cell value, and disables the label when the value is zero before saving the file.
// Keywords: Aspose.Cells | C# | .NET | stacked column chart | data labels | zero value | hide labels | chart series points | conditional label visibility | Excel automation | chart customization
// Common Searches: Aspose.Cells hide zero data labels | C# stacked column chart hide labels | remove zero value labels Aspose.Cells | conditional data label visibility .NET | how to hide chart point labels in Aspose.Cells
// Developer Intent: Programmatically suppress data labels for points with a zero value in a stacked column chart.
// Use Cases: Quarterly sales dashboard where categories with no sales are displayed without label clutter. | Financial expense chart that omits zero‑amount entries for clearer presentation. | Automated report generation that cleans up chart labels based on underlying cell values. | Dynamic Excel workbook creation where label visibility follows business rules, such as hiding labels below a threshold.
// AI Prompts: Generate C# code using Aspose.Cells to create a stacked column chart and hide data labels for points whose cell value equals zero. | Explain step‑by‑step how to iterate through series points in Aspose.Cells and toggle DataLabels.ShowValue based on the cell's numeric value. | Provide a modification of the example to hide labels for values less than a configurable threshold instead of only zero. | Show how to apply the same zero‑label hiding technique to other chart types (e.g., bar, line) with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example builds a workbook, inserts sample data that includes zero values, creates a stacked column chart, enables data labels for all series, then scans each point, reads its cell value, and disables the label when the value is zero before saving the file.
    public class HideZeroDataLabelsInStackedColumnChart
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with some zero values
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Series 1
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(0);   // zero value
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Series 2
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(0);   // zero value
            sheet.Cells["C5"].PutValue(35);

            // Add a stacked column chart
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set data ranges for the two series
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for each series
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = true;
            }

            // Hide data labels for points whose value is zero
            for (int s = 0; s < chart.NSeries.Count; s++)
            {
                Series series = chart.NSeries[s];
                for (int i = 0; i < series.Points.Count; i++)
                {
                    // Determine the cell that holds the point's value
                    // Data starts at row 2 (index 1) and column B (index 1) for the first series
                    int rowIndex = i + 1;               // zero‑based row index for data rows
                    int colIndex = 1 + s;               // zero‑based column index (B=1, C=2, …)

                    double pointValue = sheet.Cells[rowIndex, colIndex].DoubleValue;

                    // If the value is zero, hide its data label
                    if (Math.Abs(pointValue) < double.Epsilon)
                    {
                        series.Points[i].DataLabels.ShowValue = false;
                    }
                }
            }

            // Save the workbook
            workbook.Save("StackedColumnHideZeroLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}
