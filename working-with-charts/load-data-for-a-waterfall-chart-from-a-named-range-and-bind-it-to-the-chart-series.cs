// Title: How to load a named range and bind its columns to a Waterfall chart series using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that fetches the 'WaterfallData' named range and maps its first column to the category axis and second column to the value axis when constructing a Waterfall chart with Aspose.Cells. | Illustrate assigning separate category and value address strings derived from a named range to a Waterfall chart series and setting a custom chart title in Aspose.Cells. | Provide error‑handling for missing input workbook or absent named range, then save the workbook after linking the chart series to the named‑range data in C#.
// Common Searches: Aspose.Cells example for creating a waterfall chart from a predefined named range | C# code to set category and value ranges for a waterfall chart in Aspose.Cells | How to reference Excel named ranges when building charts with Aspose.Cells .NET | Assigning series data addresses to a waterfall chart using Aspose.Cells API | Saving workbook after adding waterfall chart with Aspose.Cells in C#
// Tags: waterfall chart series binding Aspose.Cells | named range chart data source C# | set series values address Aspose.Cells | waterfall chart insertion Aspose.Cells | chart title configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing workbook, retrieves the 'WaterfallData' named range, creates a Waterfall chart, binds the first column as categories and the second column as values to the chart series, sets a chart title, and saves the workbook.
class WaterfallChartExample
{
    static void Main()
    {
        try
        {
            const string inputFile = "InputData.xlsx";
            const string outputFile = "WaterfallChartOutput.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file '{inputFile}' not found.");

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the named range that contains the waterfall data
            // The named range should include both categories (first column) and values (second column)
            Aspose.Cells.Range dataRange = workbook.Worksheets.GetRangeByName("WaterfallData");
            if (dataRange == null)
                throw new InvalidOperationException("Named range 'WaterfallData' was not found.");

            // Determine addresses for categories (first column) and values (second column)
            int firstRow = dataRange.FirstRow;
            int firstCol = dataRange.FirstColumn;
            int rowCount = dataRange.RowCount;
            int lastRow = firstRow + rowCount - 1;

            // Category range (first column of the named range)
            string catAddress = $"{sheet.Name}!{CellsHelper.CellIndexToName(firstRow, firstCol)}:{CellsHelper.CellIndexToName(lastRow, firstCol)}";

            // Values range (second column of the named range)
            string valAddress = $"{sheet.Name}!{CellsHelper.CellIndexToName(firstRow, firstCol + 1)}:{CellsHelper.CellIndexToName(lastRow, firstCol + 1)}";

            // Add a Waterfall chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartTopRow = 5;
            int chartLeftColumn = 0;
            int chartBottomRow = 25;
            int chartRightColumn = 10;
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, chartTopRow, chartLeftColumn, chartBottomRow, chartRightColumn);
            Chart chart = sheet.Charts[chartIndex];

            // Add the series using the category range; for column‑type charts the second argument is a bool indicating vertical orientation
            int seriesIndex = chart.NSeries.Add(catAddress, true);
            Series series = chart.NSeries[seriesIndex];
            series.Name = "Waterfall Series";

            // Assign the values range to the series (correct property is Values)
            series.Values = valAddress;

            // Optional: set chart title
            chart.Title.Text = "Waterfall Chart from Named Range";

            // Save the workbook with the chart
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
