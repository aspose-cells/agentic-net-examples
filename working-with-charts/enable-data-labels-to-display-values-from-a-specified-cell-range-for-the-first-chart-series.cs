// Title: Add custom data‑label text from a cell range to the first series of a column chart with Aspose.Cells for .NET
// Description: This C# example creates a workbook, fills columns A‑C with categories, values, and custom label strings, inserts a column chart, links the first series to the value range, and configures the series to display data labels using the cell range C2:C4 via ShowCellRange and LinkedSource. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# chart data labels | LinkedSource | ShowCellRange | custom label text | column chart | Excel automation | set data labels from cells | Aspose.Cells for .NET | chart series label range
// Common Searches: Aspose.Cells link data labels to cell range | C# set custom chart labels using LinkedSource | ShowCellRange example Aspose.Cells | How to use cell values as chart data labels .NET | Add custom text to first series data labels Aspose
// Developer Intent: Show custom text from cells C2:C4 as data labels for the first series of a column chart.
// Use Cases: Create a sales bar chart where each column shows a formatted label like "100 units" taken from a separate column. | Build a financial report chart that updates its data‑label text automatically when the source cells in column C are edited. | Generate presentation‑ready charts with readable, pre‑formatted labels without altering the numeric series.
// AI Prompts: Write C# code with Aspose.Cells to apply custom data‑label text from range D2:D5 to the second series of a line chart. | Explain the interaction between ShowCellRange and LinkedSource for displaying custom labels in an Aspose.Cells chart. | Provide an example that changes the data‑label position based on the value magnitude for a column chart using Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, fills columns A‑C with categories, values, and custom label strings, inserts a column chart, links the first series to the value range, and configures the series to display data labels using the cell range C2:C4 via ShowCellRange and LinkedSource. The workbook is saved as an XLSX file.
    public class DataLabelsFromCellRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for categories and values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                // Put custom label texts in another column (these will be shown as data labels)
                sheet.Cells["C1"].PutValue("Label");
                sheet.Cells["C2"].PutValue("100 units");
                sheet.Cells["C3"].PutValue("200 units");
                sheet.Cells["C4"].PutValue("300 units");

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and categories
                chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

                // Access the first series
                Series firstSeries = chart.NSeries[0];

                // Enable data labels and configure them to use the cell range defined in column C
                firstSeries.DataLabels.ShowValue = true;          // Show the numeric value (optional)
                firstSeries.DataLabels.ShowCellRange = true;      // Use cell range for label text
                firstSeries.DataLabels.LinkedSource = "C2:C4";    // Range that contains custom label texts

                // Optional: set label position and font color for better visibility
                firstSeries.DataLabels.Position = LabelPositionType.InsideEnd;
                firstSeries.DataLabels.Font.Color = Color.Blue;

                // Determine output path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DataLabelsFromCellRange.xlsx");

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DataLabelsFromCellRange.Run();
        }
    }
}
