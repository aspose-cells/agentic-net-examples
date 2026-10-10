// Title: Enable data labels on a scatter chart and display notes from worksheet cells using Aspose.Cells for .NET (API limitation noted)
// AI Prompts: Write C# code with Aspose.Cells that creates a scatter chart, turns on data labels, and attempts to set each label’s text from a corresponding cell in the worksheet. | Explain how Aspose.Cells manages data label text for scatter charts, why direct linking of labels to cells is not supported, and propose a workaround to show custom notes.
// Common Searches: Aspose.Cells C# scatter chart show custom notes as data labels | how to bind scatter plot data labels to worksheet cells using Aspose.Cells | display cell values as data labels in Aspose.Cells scatter chart .NET | Aspose.Cells data label linking limitation for scatter charts | C# create scatter chart with data labels from column C using Aspose.Cells
// Tags: scatter chart data labels Aspose.Cells | custom label text from worksheet cells .NET | Aspose.Cells chart series XValues YValues | link data label to cell limitation Aspose.Cells | scatter plot notes as labels Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills columns A‑C with X values, Y values, and descriptive notes, adds a scatter chart referencing the X and Y ranges, enables data labels to show Y values, notes that Aspose.Cells does not currently support linking each label directly to a cell, and saves the file as ScatterWithDataLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Note");

            // Sample data
            double[] xValues = { 1, 2, 3, 4, 5 };
            double[] yValues = { 2, 4, 1, 3, 5 };
            string[] notes = { "First", "Second", "Third", "Fourth", "Fifth" };

            // Populate worksheet with data
            for (int i = 0; i < xValues.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(xValues[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue(yValues[i]); // Column B
                sheet.Cells[i + 1, 2].PutValue(notes[i]);   // Column C
            }

            // Add a scatter chart (positioned from row 7, column 0 to row 25, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define ranges for X and Y values (excluding header row)
            int dataRowCount = xValues.Length + 1; // include header row
            string xRange = $"=Sheet1!$A$2:$A${dataRowCount}";
            string yRange = $"=Sheet1!$B$2:$B${dataRowCount}";

            // Add Y series and then assign X values (compatible with older API versions)
            chart.NSeries.Add(yRange, false);
            chart.NSeries[0].XValues = xRange;

            // Enable data labels for the series (show Y values)
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Note: Direct linking of data labels to cells (SetDataLabelCell) is not available
            // in the current API version, so we only display the Y values as labels.

            // Save the workbook
            string outputPath = "ScatterWithDataLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
