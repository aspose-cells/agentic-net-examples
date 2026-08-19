// Title: Enable Major Tick Marks on the X‑Axis of a Column Chart with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a column chart, binds the data range, and sets the X‑axis (CategoryAxis) major tick marks to outside using TickMarkType.Outside before saving the file.
// Keywords: Aspose.Cells | .NET chart formatting | C# Excel automation | CategoryAxis major tick mark | TickMarkType Outside | X‑axis tick marks | column chart axis styling | Excel chart customization | programmatic chart axis | Aspose.Cells examples
// Common Searches: Aspose.Cells set X axis major tick marks C# | how to show tick marks on chart axis using Aspose.Cells | C# enable outside tick marks for column chart | CategoryAxis.TickMark Aspose.Cells .NET | Excel chart axis formatting with Aspose.Cells
// Developer Intent: Add visible major tick marks to the X‑axis of a chart to make category labels easier to read.
// Use Cases: Generate Excel reports with column charts that have clearly separated X‑axis labels. | Standardize axis appearance across multiple charts in a workbook for consistent documentation. | Prepare printable spreadsheets where prominent X‑axis tick marks improve visual clarity.
// AI Prompts: Provide C# code using Aspose.Cells to set major tick marks on both X and Y axes of a line chart. | Show how to change tick mark style to Inside and adjust length for a bar chart with Aspose.Cells. | Explain how to toggle major tick marks on a chart's CategoryAxis based on a configuration setting in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a column chart, binds the data range, and sets the X‑axis (CategoryAxis) major tick marks to outside using TickMarkType.Outside before saving the file.
    public class EnableMajorTickMarksOnXAxis
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

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable major tick marks on the X‑axis (category axis)
                chart.CategoryAxis.MajorTickMark = TickMarkType.Outside;

                // Save the workbook to a file
                string outputPath = "EnableMajorTickMarksOnXAxis.xlsx";
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
            EnableMajorTickMarksOnXAxis.Run();
        }
    }
}
