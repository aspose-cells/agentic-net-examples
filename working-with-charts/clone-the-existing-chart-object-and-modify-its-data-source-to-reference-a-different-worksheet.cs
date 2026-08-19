// Title: Clone a Chart and Reassign Its Data Source to a Different Worksheet with Aspose.Cells for .NET (C#)
// Description: C# code that loads or creates an Excel workbook, clones an existing chart onto a new sheet, redirects the cloned chart’s series and category ranges to a separate data worksheet, copies visual properties such as title and style, and saves the updated workbook.
// Keywords: Aspose.Cells | C# chart cloning | change chart data source | Excel chart duplicate | NSeries range Aspose | chart style copy | worksheet data source | clone chart .NET | chart data binding | Aspose.Cells example
// Common Searches: Aspose.Cells clone chart to new sheet | Change chart data range to another worksheet Aspose.Cells | Copy chart formatting and set new data source C# | How to duplicate Excel chart with Aspose.Cells | Set NSeries range on different sheet after cloning
// Developer Intent: Duplicate an existing chart and bind it to data on a different worksheet.
// Use Cases: Create a summary page that reuses the original chart layout while displaying data from a separate analysis sheet. | Generate reports where the source chart remains unchanged and a cloned chart reflects updated values from another dataset. | Automate workbook creation that applies consistent chart styling across multiple data sets by cloning and retargeting the data source.
// AI Prompts: Write C# code using Aspose.Cells to clone an existing chart and set its NSeries and CategoryData to ranges on a different worksheet. | Show how to copy a chart’s title and style to a cloned chart on another sheet while changing the data source. | Explain error‑handling steps for cloning a chart when the source workbook might not contain any charts.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# code that loads or creates an Excel workbook, clones an existing chart onto a new sheet, redirects the cloned chart’s series and category ranges to a separate data worksheet, copies visual properties such as title and style, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook or create a sample one if it does not exist.
            if (File.Exists(sourcePath))
            {
                workbook = new Workbook(sourcePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Sample data
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["A2"].PutValue("A");
                dataSheet.Cells["A3"].PutValue("B");
                dataSheet.Cells["A4"].PutValue("C");
                dataSheet.Cells["B1"].PutValue("Value");
                dataSheet.Cells["B2"].PutValue(10);
                dataSheet.Cells["B3"].PutValue(20);
                dataSheet.Cells["B4"].PutValue(30);

                // Sample chart
                int chartIdx = dataSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart sampleChart = dataSheet.Charts[chartIdx];
                sampleChart.NSeries.Add("Data!B2:B4", true);
                sampleChart.NSeries.CategoryData = "Data!A2:A4";
                sampleChart.Title.Text = "Sample Chart";

                // Save the generated source workbook for future runs.
                workbook.Save(sourcePath);
            }

            // Ensure the source worksheet contains at least one chart.
            Worksheet sourceSheet = workbook.Worksheets[0];
            if (sourceSheet.Charts.Count == 0)
                throw new InvalidOperationException("The source worksheet does not contain any charts to clone.");

            Chart sourceChart = sourceSheet.Charts[0];

            // Create a new worksheet that will hold the cloned chart.
            Worksheet chartSheet = workbook.Worksheets.Add("ClonedChartSheet");

            // Create another worksheet that will provide the new data source.
            Worksheet newDataSheet = workbook.Worksheets.Add("NewData");
            newDataSheet.Cells["A1"].PutValue("Category");
            newDataSheet.Cells["A2"].PutValue("X");
            newDataSheet.Cells["A3"].PutValue("Y");
            newDataSheet.Cells["A4"].PutValue("Z");
            newDataSheet.Cells["B1"].PutValue("Value");
            newDataSheet.Cells["B2"].PutValue(10);
            newDataSheet.Cells["B3"].PutValue(20);
            newDataSheet.Cells["B4"].PutValue(30);

            // Add a new chart on the chart sheet with the same type as the source chart.
            int clonedChartIndex = chartSheet.Charts.Add(sourceChart.Type, 5, 0, 15, 5);
            Chart clonedChart = chartSheet.Charts[clonedChartIndex];

            // Set the data source of the cloned chart to refer to the new worksheet.
            clonedChart.NSeries.Add("NewData!B2:B4", true);
            clonedChart.NSeries.CategoryData = "NewData!A2:A4";

            // Copy visual properties from the source chart.
            clonedChart.Title.Text = sourceChart.Title.Text;
            clonedChart.Style = sourceChart.Style;

            // Save the workbook with the cloned chart.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
