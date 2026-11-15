// Title: Add a multiline TextBox to a chart and position it at the plot area’s top‑right corner with Aspose.Cells for .NET
// Common Searches: how to insert a multiline textbox into a chart using Aspose.Cells C# | align a shape to the top right of a chart plot area in Aspose.Cells .NET | set textbox left and top coordinates based on chart plot area dimensions Aspose.Cells
// Tags: multiline textbox shape insertion Aspose.Cells | chart plot area coordinate retrieval Aspose.Cells | shape positioning at chart top‑right Aspose.Cells | Aspose.Cells chart shape alignment .NET | add text box to worksheet chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a workbook, adds a column chart, inserts a multiline TextBox shape on the worksheet, reads the chart’s PlotArea size and location, moves the TextBox to the plot area’s top‑right corner, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (rows 5‑25, columns 0‑10)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add a multiline TextBox shape to the worksheet.
                // Parameters: upper‑left row, upper‑left column, row offset, column offset, width (points), height (points)
                Shape textBox = sheet.Shapes.AddTextBox(5, 0, 0, 0, 200, 80);

                // Set multiline text (use \r\n for line breaks)
                textBox.Text = "First line\r\nSecond line\r\nThird line";

                // Optional formatting
                textBox.Font.IsBold = true;
                textBox.Font.Size = 12;

                // Retrieve plot area dimensions and position
                double plotAreaWidth = chart.PlotArea.Width;
                double plotAreaHeight = chart.PlotArea.Height;
                double plotAreaX = chart.PlotArea.X; // left position of plot area
                double plotAreaY = chart.PlotArea.Y; // top position of plot area

                // Align the TextBox to the top‑right corner of the plot area
                textBox.Left = (int)(plotAreaX + plotAreaWidth - textBox.Width);
                textBox.Top = (int)plotAreaY;

                // Save the workbook
                string outputFile = "ChartWithMultilineTextBox.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputFile)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
