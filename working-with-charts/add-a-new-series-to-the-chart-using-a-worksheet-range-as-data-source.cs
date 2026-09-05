// Title: Add a second data series to a column chart in Aspose.Cells using worksheet ranges (C#)
// AI Prompts: Create a new workbook, fill columns A‑C with category labels and two numeric series, add a column chart, bind the first series to B2:B5, set the category axis to A2:A5, then add a second series from C2:C5, and save the file as an .xlsx document. | Generate C# code that demonstrates how to call NSeries.Add with a worksheet range to attach an additional series to an existing Aspose.Cells column chart. | Write a snippet that shows how to assign both the series data range and the category data range for a column chart in Aspose.Cells, then export the workbook.
// Common Searches: Aspose.Cells C# add another series to column chart from cell range | how to bind multiple data series to a chart using Aspose.Cells | set category axis range for Aspose.Cells chart programmatically C# | example of NSeries.Add with worksheet range in Aspose.Cells | create column chart with two series using Aspose.Cells C#
// Tags: NSeries.Add range binding Aspose.Cells C# | column chart multiple series Aspose.Cells | set chart category data Aspose.Cells | chart data source worksheet range Aspose.Cells | save workbook with chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, populates columns A‑C with categories and two numeric series, adds a column chart, assigns the first series and category axis using cell ranges, adds a second series from another range, and saves the workbook as ChartWithMultipleSeries.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data:
        // Column A – Category labels
        // Column B – First data series
        // Column C – Second data series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");

        string[] categories = { "A", "B", "C", "D" };
        int[] series1 = { 10, 20, 30, 40 };
        int[] series2 = { 15, 25, 35, 45 };

        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(categories[i]); // A column
            sheet.Cells[i + 2, 1].PutValue(series1[i]);   // B column
            sheet.Cells[i + 2, 2].PutValue(series2[i]);   // C column
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the first series (Series1) using the range B2:B5
        chart.NSeries.Add("=Sheet1!$B$2:$B$5", true);
        // Define the category (X‑axis) data range
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

        // Add a new series (Series2) using the range C2:C5
        chart.NSeries.Add("=Sheet1!$C$2:$C$5", true);

        // Save the workbook with the chart containing multiple series
        workbook.Save("ChartWithMultipleSeries.xlsx");
    }
}
