// Title: How to reference a range on another worksheet as the data source for a column chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart series to a range on a different worksheet | Create a column chart in Aspose.Cells that uses data from another sheet | Cross‑sheet chart data source example with Aspose.Cells for .NET | Using NSeries.Add to link chart to external worksheet range in C# | Aspose.Cells chart linking to DataSheet!A1:B5 in a separate chart sheet | How to bind Excel chart to data on a separate worksheet with Aspose.Cells
// Tags: Aspose.Cells chart data from another worksheet | C# NSeries.Add external range reference | Aspose.Cells column chart using separate sheet data | Excel chart linked to DataSheet range Aspose.Cells | Aspose.Cells multi‑sheet chart linking

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCrossSheetChart
{
    // The example creates a workbook, populates a sheet named 'DataSheet' with category and value data, adds a second sheet 'ChartSheet', inserts a column chart, and binds the chart's series to the range 'DataSheet!A1:B5' (including categories) using NSeries.Add, enabling cross‑sheet chart data linking, then saves the workbook as 'CrossSheetChart.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle create rule)
                Workbook workbook = new Workbook();

                // Access the default first worksheet and add some sample data
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "DataSheet";

                // Populate data range A1:B5
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Value");
                dataSheet.Cells["A2"].PutValue("A");
                dataSheet.Cells["A3"].PutValue("B");
                dataSheet.Cells["A4"].PutValue("C");
                dataSheet.Cells["A5"].PutValue("D");
                dataSheet.Cells["B2"].PutValue(10);
                dataSheet.Cells["B3"].PutValue(20);
                dataSheet.Cells["B4"].PutValue(30);
                dataSheet.Cells["B5"].PutValue(40);

                // Add a new worksheet that will contain the chart
                Worksheet chartSheet = workbook.Worksheets.Add("ChartSheet");

                // Add a column chart to the chart sheet
                int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = chartSheet.Charts[chartIndex];

                // Reference the data range on a different worksheet (cross‑sheet linking)
                // The range "DataSheet!A1:B5" includes categories (A column) and values (B column)
                // The second parameter indicates that the first row contains category names
                chart.NSeries.Add("DataSheet!A1:B5", true);

                // Set the chart title (optional)
                chart.Title.Text = "Cross‑Sheet Column Chart";

                // Determine output file path
                string outputPath = "CrossSheetChart.xlsx";

                // Save the workbook (lifecycle save rule)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
