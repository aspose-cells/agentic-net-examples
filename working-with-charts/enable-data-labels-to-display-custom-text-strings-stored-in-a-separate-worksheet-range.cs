// Title: Aspose.Cells C# – Show custom chart data labels from another worksheet range
// Description: Creates a workbook with a data sheet and a separate label sheet, adds a column chart, and configures the series to display custom text from the range Labels!A2:A4 instead of the default values, applying optional font color and position before saving the file.
// Keywords: Aspose.Cells | C# | custom data labels | chart label range | LinkedSource property | Excel chart styling | cell range labels | column chart Aspose.Cells | .NET chart data labels | display custom text in chart
// Common Searches: Aspose.Cells set custom data label text from cell range | C# chart data labels linked to another worksheet | How to use LinkedSource for chart labels in Aspose.Cells | Display custom strings as data labels in Aspose.Cells chart | Aspose.Cells hide default values and show custom labels
// Developer Intent: Use a separate worksheet range to supply custom text for chart data labels in Aspose.Cells.
// Use Cases: Show descriptive labels (e.g., product names) on a column chart instead of numeric values. | Maintain label text in a dedicated sheet so updates automatically reflect in the chart. | Apply specific font color or position to custom labels for better visual emphasis.
// AI Prompts: Write C# code with Aspose.Cells that links chart data labels to a cell range on a different worksheet and hides the default values. | Show how to set font color and label position for custom data labels using the LinkedSource property. | Explain how to adapt the example for a line chart while keeping custom labels from another sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook with a data sheet and a separate label sheet, adds a column chart, and configures the series to display custom text from the range Labels!A2:A4 instead of the default values, applying optional font color and position before saving the file.
    public class CustomDataLabelsFromRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Worksheet 1 – source data for the chart
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Category (X) values
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["A2"].PutValue("A");
                dataSheet.Cells["A3"].PutValue("B");
                dataSheet.Cells["A4"].PutValue("C");

                // Numeric (Y) values
                dataSheet.Cells["B1"].PutValue("Value");
                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["B3"].PutValue(85);
                dataSheet.Cells["B4"].PutValue(65);

                // -------------------------------------------------
                // Worksheet 2 – custom label strings
                // -------------------------------------------------
                Worksheet labelSheet = workbook.Worksheets.Add("Labels");

                // Custom text for each data point
                labelSheet.Cells["A1"].PutValue("CustomLabel");
                labelSheet.Cells["A2"].PutValue("First");
                labelSheet.Cells["A3"].PutValue("Second");
                labelSheet.Cells["A4"].PutValue("Third");

                // -------------------------------------------------
                // Add a column chart to the data sheet
                // -------------------------------------------------
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 6, 0, 22, 12);
                Chart chart = dataSheet.Charts[chartIndex];

                // Bind the series to the numeric values and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // Configure data labels to use the custom text range
                // -------------------------------------------------
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = false;               // hide default value
                series.DataLabels.ShowCellRange = true;            // enable cell range display
                series.DataLabels.LinkedSource = "Labels!A2:A4";    // link to custom strings
                series.DataLabels.Font.Color = Color.Blue;         // optional styling
                series.DataLabels.Position = LabelPositionType.InsideEnd; // optional position

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("CustomDataLabelsFromRange.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CustomDataLabelsFromRange.Run();
        }
    }
}
