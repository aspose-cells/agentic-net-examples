// Title: Create a line chart from cells A1‑A10 in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to add a line chart using data range A1:A10 | how to use NSeries.Add to plot column A values as a line chart in Aspose.Cells | generate an XLSX file with a line chart from cells A1 to A10 using Aspose.Cells .NET | Aspose.Cells line chart positioning on worksheet C# code sample | set chart title and series name for a line chart in Aspose.Cells C#
// Tags: Aspose.Cells line chart from cell range | C# add NSeries range A1:A10 | chart positioning with Aspose.Cells worksheet | set chart title and series name Aspose.Cells | save workbook with chart as XLSX using Aspose.Cells | generate Excel line chart programmatically .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills cells A1 through A10 with numeric values, adds a line chart to the first worksheet, sets the chart title to "Line Chart", adds a series referencing the range A1:A10 with the series name "Sample Data", positions the chart on the sheet, and saves the file as LineChart.xlsx while handling potential exceptions.
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

            // Populate cells A1 through A10 with sample data
            for (int i = 0; i < 10; i++)
            {
                // Cells are 0‑based index: row i, column 0 corresponds to A(i+1)
                sheet.Cells[i, 0].PutValue(i + 1); // Example values 1..10
            }

            // Add a line chart to the worksheet
            // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
            int chartIndex = sheet.Charts.Add(ChartType.Line, 12, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Line Chart";

            // Add a series using data from A1:A10
            // The second parameter (true) indicates that the data is plotted against the default category axis
            chart.NSeries.Add("A1:A10", true);

            // Optionally, set the series name (displayed in legend)
            chart.NSeries[0].Name = "Sample Data";

            // Save the workbook to a file
            string outputPath = "LineChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
