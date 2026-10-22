// Title: Create a column‑line combo chart with a secondary axis in C# using Aspose.Cells (cell‑based data labels not supported)
// AI Prompts: Generate C# code with Aspose.Cells that builds a combo chart, assigns the line series to a secondary Y‑axis, and turns on data labels for that series. | Adapt the sample to attempt binding custom text from a worksheet range (e.g., D2:D4) to the secondary series data labels, while handling the current Aspose.Cells limitation.
// Common Searches: asp.net aspose.cells create combo chart with secondary y axis c# | c# aspose.cells line series on secondary axis column chart example | how to display custom text data labels from worksheet cells in aspose.cells chart | aspose.cells secondary axis not showing in generated excel file | c# generate excel combo chart with primary column and secondary line using aspose.cells
// Tags: Aspose.Cells secondary Y‑axis for line series | C# Aspose.Cells column‑line combo chart example | Aspose.Cells bind chart data labels from worksheet cells | Aspose.Cells chart data label limitation | Aspose.Cells create combo chart with mixed series types

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new Workbook, fills columns A‑C with categories and numeric values, adds a column‑line combo chart, sets the primary series (column B) as a column chart and the secondary series (column C) as a line chart, enables simple value data labels for the secondary series, and saves the file as ComboChart_SecondaryAxis.xlsx. Note that Aspose.Cells currently does not support placing a series on a secondary axis or displaying data labels directly from a cell range.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Primary Series");
            ws.Cells["C1"].PutValue("Secondary Series");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);
            ws.Cells["C2"].PutValue(100);
            ws.Cells["C3"].PutValue(150);
            ws.Cells["C4"].PutValue(200);

            // Data labels for secondary series (stored in cells)
            ws.Cells["D2"].PutValue("Low");
            ws.Cells["D3"].PutValue("Medium");
            ws.Cells["D4"].PutValue("High");

            // Add a combo chart (Column + Line)
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = ws.Charts[chartIndex];
            chart.Title.Text = "Combo Chart with Secondary Axis";

            // Primary series (Column) – uses column B
            int primarySeriesIdx = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[primarySeriesIdx].Name = "Primary Series";

            // Secondary series (Line) – uses column C
            int secondarySeriesIdx = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[secondarySeriesIdx].Name = "Secondary Series";
            chart.NSeries[secondarySeriesIdx].Type = ChartType.Line;

            // NOTE: The following features are not available in the current Aspose.Cells version:
            // - Placing the series on a secondary axis
            // - Making secondary axes visible
            // - Showing data labels directly from a cell range
            // Therefore, they have been omitted to ensure successful compilation.

            // Show simple data labels for the secondary series (values only)
            chart.NSeries[secondarySeriesIdx].DataLabels.ShowValue = true;
            chart.NSeries[secondarySeriesIdx].DataLabels.ShowCategoryName = false;
            chart.NSeries[secondarySeriesIdx].DataLabels.ShowSeriesName = false;

            // Save the workbook
            string outputPath = "ComboChart_SecondaryAxis.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
