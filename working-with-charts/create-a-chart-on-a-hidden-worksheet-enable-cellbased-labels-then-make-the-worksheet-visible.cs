// Title: Generate a column chart with cell‑based data labels on a hidden worksheet and then reveal the sheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add column chart to hidden worksheet with cell based labels | How to enable linked source data labels for a chart on a hidden sheet in Aspose.Cells | Unhide a worksheet after creating a chart that uses custom label range with Aspose.Cells .NET | Set DataLabels.ShowCellRange true for chart series in Aspose.Cells C# example | Create chart on a hidden sheet then make the sheet visible using Aspose.Cells
// Tags: column chart with cell‑based labels Aspose.Cells | hidden worksheet chart generation .NET | linked source data labels Aspose.Cells | unhide worksheet after chart creation Aspose.Cells | Aspose.Cells chart on hidden sheet example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsChartOnHiddenSheet
{
    // The sample creates a new workbook, adds a hidden worksheet, populates it with category, value, and custom label data, inserts a column chart, configures the series to display cell‑based data labels linked to a label range, makes the worksheet visible again, and saves the file as ChartOnHiddenSheet.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a new worksheet that will be hidden initially
                int hiddenSheetIndex = workbook.Worksheets.Add();
                Worksheet hiddenSheet = workbook.Worksheets[hiddenSheetIndex];
                hiddenSheet.Name = "HiddenChartSheet";

                // Hide the worksheet
                hiddenSheet.IsVisible = false;

                // Populate sample data
                // Column A – Category
                // Column B – Value
                // Column C – Custom label (cell‑based label)
                hiddenSheet.Cells["A1"].PutValue("Category");
                hiddenSheet.Cells["B1"].PutValue("Value");
                hiddenSheet.Cells["C1"].PutValue("Label");

                hiddenSheet.Cells["A2"].PutValue("Apple");
                hiddenSheet.Cells["B2"].PutValue(120);
                hiddenSheet.Cells["C2"].PutValue("120 units");

                hiddenSheet.Cells["A3"].PutValue("Orange");
                hiddenSheet.Cells["B3"].PutValue(80);
                hiddenSheet.Cells["C3"].PutValue("80 units");

                // Add a column chart on the hidden worksheet
                int chartIndex = hiddenSheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = hiddenSheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B3", true);          // Values
                chart.NSeries.CategoryData = "A2:A3";      // Categories

                // Enable cell‑based data labels
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;        // Show the numeric value (optional)
                series.DataLabels.ShowCellRange = true;    // Enable cell‑range labels
                series.DataLabels.LinkedSource = "C2:C3";  // Range that contains the custom labels
                series.DataLabels.Font.Color = Color.Blue; // Optional styling

                // Optionally give the chart a title
                chart.Title.Text = "Fruit Sales (Cell‑Based Labels)";

                // Make the hidden worksheet visible again
                hiddenSheet.IsVisible = true;

                // Save the workbook
                string outputPath = "ChartOnHiddenSheet.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
