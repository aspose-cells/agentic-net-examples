using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            try
            {
                InsertMultilineTextboxInChart.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    class InsertMultilineTextboxInChart
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
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Add a multiline textbox to the chart positioned at the plot area's top‑right.
                // Units are 1/4000 of the chart area. Left is set near the right edge (e.g., 3000).
                TextBox txtBox = chart.Shapes.AddTextBoxInChart(
                    top: 0,          // top offset
                    left: 3000,      // left offset (near right edge)
                    height: 200,     // height of the textbox
                    width: 400);     // width of the textbox

                // Set multiline text using line breaks
                txtBox.Text = "First line\nSecond line\nThird line";

                // Allow text to overflow if the content exceeds the box size
                txtBox.TextBoxOptions.AllowTextToOverflow = true;

                // Optional formatting
                txtBox.Font.Color = Color.DarkBlue;
                txtBox.Font.Size = 12;
                txtBox.Font.IsBold = true;

                // Save the workbook
                string outputPath = "MultilineTextboxInChart.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Capture any errors that occur during workbook manipulation
                Console.WriteLine($"Error in InsertMultilineTextboxInChart: {ex.Message}");
                throw;
            }
        }
    }
}