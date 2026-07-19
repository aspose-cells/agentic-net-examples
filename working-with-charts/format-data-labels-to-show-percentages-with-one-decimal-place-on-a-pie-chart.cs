// Title: C# – Aspose.Cells: Show Pie‑Chart Data Labels as 0.0% Percentages
// Description: Creates an Excel workbook, adds category/value rows, inserts a pie chart, and configures the first series to display data‑label percentages with one‑decimal precision (0.0%) while hiding raw values, then saves the file.
// Keywords: Aspose.Cells | C# | .NET | pie chart | data labels | percentage format | one decimal place | Excel export | chart formatting | NumberFormat | ShowPercentage
// Common Searches: Aspose.Cells pie chart show percentage one decimal | C# set data label format 0.0% Aspose.Cells | hide values display only percentages chart Aspose.Cells .NET | format pie chart labels as percentage Aspose.Cells example | Excel pie chart data labels custom number format C#
// Developer Intent: Format pie‑chart labels to show percentages with one decimal place.
// Use Cases: Generate a sales‑distribution workbook where each slice is labeled only with a 0.0% value. | Create a dashboard Excel file that presents market‑share data using a clean percentage‑only pie chart. | Automate reporting for finance teams by exporting Excel charts with rounded percentage labels.
// AI Prompts: Write C# code using Aspose.Cells to add a pie chart and set its data labels to 0.0% format. | Show how to hide raw values and display only one‑decimal‑place percentages on a pie chart with Aspose.Cells .NET. | Explain the steps to apply a custom NumberFormat to chart data labels in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates an Excel workbook, adds category/value rows, inserts a pie chart, and configures the first series to display data‑label percentages with one‑decimal precision (0.0%) while hiding raw values, then saves the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                PieChartDataLabelsPercentage.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class PieChartDataLabelsPercentage
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and configure them to show percentages
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;          // Show percentage values
            dataLabels.ShowValue = false;              // Hide raw values (optional)
            dataLabels.NumberFormat = "0.0%";          // One decimal place percentage format

            // Save the workbook to a file
            workbook.Save("PieChartDataLabelsPercentage.xlsx");
        }
    }
}
