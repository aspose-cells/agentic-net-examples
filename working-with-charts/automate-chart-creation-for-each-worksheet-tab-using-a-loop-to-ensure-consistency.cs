// Title: Create a Column Chart on Every Worksheet with Aspose.Cells for .NET (C# Loop)
// Description: This C# example builds a new workbook, adds three worksheets, fills each with sample category/value data, and uses a foreach loop to insert a column chart (rows 8‑20, columns 0‑8) that references B2:B6 for values and A2:A6 for categories. The chart title is set to the worksheet name, the layout is recalculated, and the workbook is saved as ChartsPerWorksheet.xlsx.
// Keywords: Aspose.Cells | C# | .NET | create chart programmatically | loop through worksheets | column chart | multiple worksheets | chart automation | dynamic chart title | save workbook
// Common Searches: Aspose.Cells add chart to each worksheet C# | C# loop create column chart Aspose.Cells | generate charts for multiple sheets using Aspose.Cells | dynamic chart title based on worksheet name Aspose.Cells | automate chart creation in Excel with Aspose.Cells .NET
// Developer Intent: Programmatically add identical column charts to all worksheets in a workbook using a loop.
// Use Cases: Produce monthly sales dashboards where each sheet gets a consistent column chart with a period‑specific title. | Automate template‑driven reports that require the same chart layout across dozens of worksheets. | Build a multi‑sheet financial model that updates charts automatically when new sheets are added.
// AI Prompts: Show how to modify the loop to generate a pie chart instead of a column chart for each worksheet. | Provide code to export each created chart as a PNG image while keeping the chart in the workbook. | Explain how to assign different chart types based on the worksheet name using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutomation
{
    // This C# example builds a new workbook, adds three worksheets, fills each with sample category/value data, and uses a foreach loop to insert a column chart (rows 8‑20, columns 0‑8) that references B2:B6 for values and A2:A6 for categories. The chart title is set to the worksheet name, the layout is recalculated, and the workbook is saved as ChartsPerWorksheet.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a few worksheets for demonstration
            Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
            sheet1.Name = "Sales_Q1";

            Worksheet sheet2 = workbook.Worksheets.Add("Sales_Q2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sales_Q3");

            // Array of worksheets to process
            Worksheet[] sheets = new Worksheet[] { sheet1, sheet2, sheet3 };

            // Loop through each worksheet and create a chart
            foreach (Worksheet sheet in sheets)
            {
                // Populate sample data (Category in column A, Value in column B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue((i - 1) * 10); // sample numeric values
                }

                // Add a column chart to the worksheet
                // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
                int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Define the data range for the chart
                // Series values (B2:B6) and category labels (A2:A6)
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // Optional: set a title that reflects the worksheet name
                chart.Title.Text = $"Data Chart - {sheet.Name}";

                // Recalculate the chart layout before saving
                chart.Calculate();
            }

            // Save the workbook with all charts
            workbook.Save("ChartsPerWorksheet.xlsx", SaveFormat.Xlsx);
        }
    }
}
