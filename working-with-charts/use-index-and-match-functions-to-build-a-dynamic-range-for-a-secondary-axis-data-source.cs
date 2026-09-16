// Title: Add a column chart with a dynamic secondary series using INDEX/MATCH formulas in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, populates month, primary, and secondary data, adds a column chart, and defines the secondary series range with an INDEX/MATCH formula based on start‑lookup and end‑lookup month values. | Show how to programmatically assign a dynamic range to a chart series in Aspose.Cells by building an INDEX‑MATCH expression for the secondary axis.
// Common Searches: how to use INDEX and MATCH in Aspose.Cells chart series formula | c# Aspose.Cells dynamic range for secondary axis chart | set chart series formula programmatically Aspose.Cells .NET | create column chart with dynamic secondary data source using Aspose.Cells | Aspose.Cells chart NSeries dynamic range based on lookup values
// Tags: Aspose.Cells chart series formula | dynamic range INDEX MATCH Aspose.Cells | C# secondary axis data source | column chart NSeries dynamic range | Aspose.Cells workbook chart creation .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicRangeExample
{
    // The example creates a new workbook, fills three columns with month labels, primary values, and secondary values, adds a column chart, assigns a static range to the primary series, and uses an INDEX/MATCH formula—driven by startLookup and endLookup strings—to define a dynamic range for the secondary series. The chart's value axis title is set and the workbook is saved as DynamicRangeChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data
                string[] categories = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct" };
                double[] primaryValues = { 10, 12, 15, 13, 18, 20, 22, 19, 17, 21 };
                double[] secondaryValues = { 100, 110, 115, 108, 120, 130, 125, 118, 122, 135 };

                for (int i = 0; i < categories.Length; i++)
                {
                    cells[i + 1, 0].PutValue(categories[i]);          // Column A (categories)
                    cells[i + 1, 1].PutValue(primaryValues[i]);      // Column B (primary values)
                    cells[i + 1, 2].PutValue(secondaryValues[i]);    // Column C (secondary values)
                }

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 4, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Primary series (static range)
                chart.NSeries.Add("=Sheet1!$B$2:$B$11", true);

                // Secondary series (dynamic range using INDEX & MATCH)
                string startLookup = "Mar";
                string endLookup = "Sep";

                string dynamicRangeFormula = string.Format(
                    "INDEX(Sheet1!$C$2:$C$11, MATCH(\"{0}\", Sheet1!$A$2:$A$11,0)):" +
                    "INDEX(Sheet1!$C$2:$C$11, MATCH(\"{1}\", Sheet1!$A$2:$A$11,0))",
                    startLookup, endLookup);

                chart.NSeries.Add(dynamicRangeFormula, true);

                // Set chart axis title
                chart.ValueAxis.Title.Text = "Primary Values";

                // Save the workbook
                workbook.Save("DynamicRangeChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
