// Title: Add a Column Chart to Every Worksheet in a C# Aspose.Cells Workbook Using a Loop
// Description: Creates a workbook with multiple sheets, fills each with identical category/value data, then iterates through all worksheets to insert a column chart (rows 5‑20, columns 0‑8), sets a title from the sheet name, recalculates the layout, and saves as ChartsPerWorksheet.xlsx.
// Keywords: Aspose.Cells chart loop C# | add chart to each worksheet | .NET column chart automation | multiple sheet chart generation | Aspose.Cells workbook example | dynamic chart title Aspose | C# Excel chart API | programmatic chart creation
// Common Searches: C# loop to add charts to all worksheets Aspose.Cells | how to create column chart per sheet using Aspose.Cells | set chart title from worksheet name Aspose .NET | populate identical data across sheets and generate charts
// Developer Intent: Automatically generate a column chart on every worksheet by looping through the workbook with Aspose.Cells in C#.
// Use Cases: Standardize visual reports for regional sales sheets. | Produce consistent financial charts across multiple data tabs. | Prepare template workbooks with pre‑built charts before distribution.
// AI Prompts: Generate C# code that adds a line chart to each worksheet in an existing Aspose.Cells workbook, using the same data range. | Refactor the loop to create pie charts with custom colors for every sheet and include a legend. | Explain how to configure axis formatting and legend placement for charts created in a loop with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutomation
{
    // Creates a workbook with multiple sheets, fills each with identical category/value data, then iterates through all worksheets to insert a column chart (rows 5‑20, columns 0‑8), sets a title from the sheet name, recalculates the layout, and saves as ChartsPerWorksheet.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Ensure there are multiple worksheets to demonstrate the loop
            // Add two additional worksheets (Workbook starts with one by default)
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Populate each worksheet with identical sample data
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Sample rows
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["B5"].PutValue(40);
            }

            // Loop through each worksheet and add a chart
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add a column chart positioned from row 5, column 0 to row 20, column 8
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values) and categories
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";      // Categories

                // Optional: give the chart a title that reflects the worksheet name
                chart.Title.Text = $"Sample Chart - {sheet.Name}";

                // Recalculate the chart layout before saving
                chart.Calculate();
            }

            // Save the workbook with all charts
            workbook.Save("ChartsPerWorksheet.xlsx", SaveFormat.Xlsx);
        }
    }
}
