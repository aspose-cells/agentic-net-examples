// Title: Aspose.Cells C# – Display custom chart data labels from a separate worksheet range
// Description: Creates a workbook with a data sheet and a label sheet, adds a column chart, hides the default values, links the series data labels to the range "Labels!A2:A4", applies dark‑blue font (size 10) and InsideEnd positioning, then saves the file as CustomDataLabelsFromRange.xlsx.
// Keywords: Aspose.Cells chart custom labels | C# Excel chart data labels from range | Aspose.Cells linked source labels | column chart label formatting Aspose.Cells | .NET chart data label cell range
// Common Searches: Aspose.Cells set chart data labels from another sheet | C# chart label text from cell range Aspose.Cells | How to hide values and show custom labels in Aspose.Cells chart | Format Aspose.Cells chart data label font and position
// Developer Intent: Show custom text stored on a different worksheet as data labels on a chart.
// Use Cases: Display descriptive labels (e.g., "First", "Second") instead of numeric values on a column chart. | Maintain a single source of label text that can be reused across multiple charts or series. | Apply consistent styling (color, size, position) to custom labels for better visual clarity.
// AI Prompts: Generate C# Aspose.Cells code that links chart data labels to a cell range on another worksheet and formats the labels. | Explain how to hide default data values and show custom label text from a separate sheet in an Aspose.Cells chart. | Provide error‑handling recommendations when the linked label range is missing, empty, or shorter than the data series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Creates a workbook with a data sheet and a label sheet, adds a column chart, hides the default values, links the series data labels to the range "Labels!A2:A4", applies dark‑blue font (size 10) and InsideEnd positioning, then saves the file as CustomDataLabelsFromRange.xlsx.
    public class CustomDataLabelsFromRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Sheet 1 – source data for the chart
                // -------------------------------------------------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Category (X) values
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["A2"].PutValue("Alpha");
                dataSheet.Cells["A3"].PutValue("Beta");
                dataSheet.Cells["A4"].PutValue("Gamma");

                // Numeric (Y) values
                dataSheet.Cells["B1"].PutValue("Value");
                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["B3"].PutValue(85);
                dataSheet.Cells["B4"].PutValue(150);

                // -------------------------------------------------
                // Sheet 2 – custom label texts
                // -------------------------------------------------
                Worksheet labelSheet = workbook.Worksheets.Add("Labels");
                labelSheet.Cells["A1"].PutValue("CustomLabel");
                labelSheet.Cells["A2"].PutValue("First");
                labelSheet.Cells["A3"].PutValue("Second");
                labelSheet.Cells["A4"].PutValue("Third");

                // -------------------------------------------------
                // Add a column chart to the data sheet
                // -------------------------------------------------
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = dataSheet.Charts[chartIndex];

                // Bind the series to the numeric data
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // Configure data labels to use the custom range
                // -------------------------------------------------
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = false;               // hide default value
                series.DataLabels.ShowCellRange = true;            // enable cell‑range based labels
                series.DataLabels.LinkedSource = "Labels!A2:A4";    // range with custom texts

                // Optional: format appearance of the labels
                series.DataLabels.Font.Color = Color.DarkBlue;
                series.DataLabels.Font.Size = 10;
                series.DataLabels.Position = LabelPositionType.InsideEnd;

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("CustomDataLabelsFromRange.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CustomDataLabelsFromRange.Run();
        }
    }
}
