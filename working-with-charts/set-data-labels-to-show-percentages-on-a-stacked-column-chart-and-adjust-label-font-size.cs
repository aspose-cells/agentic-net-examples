// Title: Display percentage data labels with custom font on a stacked column chart – Aspose.Cells for .NET
// Description: Creates a workbook, fills it with quarterly product figures, adds a stacked column chart, and configures each series to show only percentages on data labels while applying a 12‑pt bold blue font. The workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells | C# | stacked column chart | data labels | show percentage | custom font size | chart formatting | Excel automation | .NET chart API
// Common Searches: Aspose.Cells show percentage on stacked column chart | C# set chart data label font size Aspose.Cells | hide data label values Aspose.Cells | customize chart label color bold | add stacked column chart programmatically .NET
// Developer Intent: Add a stacked column chart and configure its data labels to display percentages only, using a custom font size and style.
// Use Cases: Quarterly sales report where each product’s share is shown as a percentage on a stacked column chart. | Financial dashboard that highlights expense composition with bold blue percentage labels while suppressing raw numbers. | Automated marketing analytics workbook that visualizes channel contributions using styled percentage labels.
// AI Prompts: Generate C# code with Aspose.Cells that creates a stacked column chart and sets data labels to show percentages only, using a 14‑pt red italic font. | Show how to modify an existing Aspose.Cells chart to hide values and display percentage labels with a custom font. | Explain the steps to apply bold blue font to percentage data labels on each series of a stacked column chart in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsStackedColumnDataLabels
{
    // Creates a workbook, fills it with quarterly product figures, adds a stacked column chart, and configures each series to show only percentages on data labels while applying a 12‑pt bold blue font. The workbook is saved as an .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a stacked column chart
                // Categories
                sheet.Cells["A1"].PutValue("Quarter");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series 1
                sheet.Cells["B1"].PutValue("Product A");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(40);
                sheet.Cells["B4"].PutValue(20);
                sheet.Cells["B5"].PutValue(10);

                // Series 2
                sheet.Cells["C1"].PutValue("Product B");
                sheet.Cells["C2"].PutValue(20);
                sheet.Cells["C3"].PutValue(30);
                sheet.Cells["C4"].PutValue(25);
                sheet.Cells["C5"].PutValue(15);

                // Series 3
                sheet.Cells["D1"].PutValue("Product C");
                sheet.Cells["D2"].PutValue(10);
                sheet.Cells["D3"].PutValue(20);
                sheet.Cells["D4"].PutValue(30);
                sheet.Cells["D5"].PutValue(25);

                // Add a stacked column chart (use ColumnStacked enum)
                int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (including all series)
                chart.NSeries.Add("B2:D5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for each series and configure them
                foreach (Series series in chart.NSeries)
                {
                    // Show percentage values on data labels
                    series.DataLabels.ShowPercentage = true;

                    // Hide the raw value if only percentages are desired
                    series.DataLabels.ShowValue = false;

                    // Adjust the font size and style of the data labels
                    series.DataLabels.Font.Size = 12;
                    series.DataLabels.Font.Color = Color.Blue;
                    series.DataLabels.Font.IsBold = true;

                    // Apply the font settings to all child label nodes
                    series.DataLabels.ApplyFont();
                }

                // Save the workbook to a file
                string outputPath = "StackedColumn_PercentageDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
