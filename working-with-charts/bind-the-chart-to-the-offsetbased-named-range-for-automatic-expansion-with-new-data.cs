// Title: Bind a Column Chart to an OFFSET‑Based Named Range for Auto‑Expanding Data in Aspose.Cells for .NET (C#)
// Description: Shows how to build a workbook, define a dynamic named range with the OFFSET formula, and attach a column chart’s series to that range so the chart grows automatically as new rows are added. Includes X‑axis categories, a chart title, and saves the file as DynamicChart.xlsx.
// Keywords: Aspose.Cells | C# | dynamic named range | OFFSET formula | auto expanding chart | column chart | chart data source | named range binding | Excel automation | programmatic chart
// Common Searches: Aspose.Cells bind chart to OFFSET named range | dynamic chart that expands with new rows .NET | use OFFSET formula for chart series in Aspose.Cells | set chart series values to a named range C# | auto‑update Excel chart with Aspose.Cells
// Developer Intent: Create a column chart whose series references an OFFSET‑based named range, allowing the chart to include additional rows without code changes.
// Use Cases: Monthly sales dashboard that updates the chart as new month data is appended. | IoT sensor monitoring where incoming measurements automatically extend the visualization. | Financial expense tracker that reflects added line items instantly in a column chart.
// AI Prompts: Generate C# code with Aspose.Cells that defines a dynamic OFFSET named range and binds a column chart’s Y‑values to it. | Explain how to make the X‑axis categories dynamic using an OFFSET range in the same example. | Provide steps to customize the chart’s title, colors, and axis labels after linking it to a dynamic named range.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChart
{
    // Shows how to build a workbook, define a dynamic named range with the OFFSET formula, and attach a column chart’s series to that range so the chart grows automatically as new rows are added. Includes X‑axis categories, a chart title, and saves the file as DynamicChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate initial data (header + 5 rows)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
                }

                // -----------------------------------------------------------------
                // Create a dynamic named range using OFFSET.
                // The range will expand automatically as new rows are added to column B.
                // Formula: =OFFSET(Sheet1!$B$2,0,0,COUNTA(Sheet1!$B:$B)-1,1)
                //   - Starts at B2 (first data cell, excluding header)
                //   - Height is the count of non‑empty cells in column B minus the header
                // -----------------------------------------------------------------
                int nameIndex = workbook.Worksheets.Names.Add("DynamicData");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                // Set the OFFSET formula (A1‑style, not R1C1, not locale‑specific)
                dynamicName.SetRefersTo(
                    $"=OFFSET({sheet.Name}!$B$2,0,0,COUNTA({sheet.Name}!$B:$B)-1,1)",
                    false,
                    false);

                // -----------------------------------------------------------------
                // Add a column chart and bind it to the dynamic named range.
                // The series uses the named range for Y‑values and the header row for X‑axis labels.
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 7);
                Chart chart = sheet.Charts[chartIndex];

                // X‑axis (categories) – static range covering enough rows
                chart.NSeries.Add($"={sheet.Name}!$A$2:$A$100", true);
                // Y‑values bound to the dynamic named range
                chart.NSeries[0].Values = $"=DynamicData";

                // Optional: set chart title
                chart.Title.Text = "Dynamic Data Chart";

                // Save the workbook
                workbook.Save("DynamicChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
