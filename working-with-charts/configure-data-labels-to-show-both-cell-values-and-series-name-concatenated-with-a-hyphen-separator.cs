// Title: Aspose.Cells C# – Show Series Name and Value in Chart Data Labels with a Hyphen Separator
// Description: Creates a workbook, adds a column chart, and configures the first series' data labels to display both the series name and the cell value. Uses a custom " - " separator and positions the labels inside the base of each column before saving the file.
// Keywords: Aspose.Cells | C# chart data labels | show series name | show value | custom separator | hyphen separator | DataLabelsSeparatorType.Custom | column chart labels | inside base label position | Excel automation example
// Common Searches: Aspose.Cells display series name and value in chart labels | C# set custom separator for chart data labels Aspose.Cells | show both value and series name in column chart Aspose.Cells | data label hyphen separator Aspose.Cells | Aspose.Cells DataLabelsSeparatorType example
// Developer Intent: Configure a chart so each data label concatenates the series name and the cell value with a hyphen.
// Use Cases: Financial reports where column labels read "SeriesName - Value" for quick comparison. | Corporate dashboards that require a branded separator in chart data labels. | Updating existing multi‑series charts to combine series names and values for clearer visualization.
// AI Prompts: Generate C# code using Aspose.Cells to add a line chart and set its data labels to show series name, value, and a '/' separator. | Explain the purpose of DataLabelsSeparatorType.Custom in Aspose.Cells and how to modify the separator string at runtime. | Provide step‑by‑step instructions to add the category name to chart data labels alongside series name and value in an existing workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart, and configures the first series' data labels to display both the series name and the cell value. Uses a custom " - " separator and positions the labels inside the base of each column before saving the file.
    public class DataLabelsSeriesNameAndValueDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the DataLabels of the first series
                DataLabels dataLabels = chart.NSeries[0].DataLabels;

                // Show both the cell value and the series name
                dataLabels.ShowValue = true;
                dataLabels.ShowSeriesName = true;

                // Use a custom separator (hyphen) between the two parts
                dataLabels.SeparatorType = DataLabelsSeparatorType.Custom;
                dataLabels.SeparatorValue = " - ";

                // Optional: set label position
                dataLabels.Position = LabelPositionType.InsideBase;

                // Save the workbook
                string outputPath = "DataLabelsSeriesNameAndValueDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            DataLabelsSeriesNameAndValueDemo.Run();
        }
    }
}
