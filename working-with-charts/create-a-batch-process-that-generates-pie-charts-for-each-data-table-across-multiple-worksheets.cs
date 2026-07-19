// Title: Batch Create Pie Charts for All Worksheets with Aspose.Cells for .NET
// Description: C# example that loads or creates a workbook, scans each worksheet, determines the data range in columns A‑B, adds a pie chart below the table, sets series and category data, applies a title and legend, and saves the file as BatchPieCharts.xlsx.
// Keywords: Aspose.Cells pie chart programmatically | C# generate chart for each worksheet | batch chart creation Aspose.Cells | loop worksheets add pie chart .NET | set chart data range Aspose.Cells | automate Excel chart generation
// Common Searches: add pie chart to every sheet Aspose.Cells | loop through worksheets create charts C# | determine last used row Aspose.Cells | position chart below data table Aspose.Cells | batch generate Excel charts .NET
// Developer Intent: Automate the addition of a pie chart to each worksheet based on its A‑B data table.
// Use Cases: Create quarterly sales pie charts for multiple region sheets in a financial report. | Produce visual summaries of product performance across months in a single workbook. | Generate charts for dynamically imported worksheets during data‑processing pipelines.
// AI Prompts: Write C# code using Aspose.Cells that iterates over all worksheets and adds a pie chart for the data in columns A and B, placing the chart below the table. | Show how to find the last used row in each worksheet and set the pie chart's value and category ranges with Aspose.Cells. | Explain how to customize the size, style, and colors of pie charts generated in a batch process with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsBatchPieCharts
{
    // C# example that loads or creates a workbook, scans each worksheet, determines the data range in columns A‑B, adds a pie chart below the table, sets series and category data, applies a title and legend, and saves the file as BatchPieCharts.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // lifecycle: create

            // Example: add two worksheets with sample data tables
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sales_Q1";
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Amount");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(1200);
            sheet1.Cells["A3"].PutValue("Banana");
            sheet1.Cells["B3"].PutValue(850);
            sheet1.Cells["A4"].PutValue("Cherry");
            sheet1.Cells["B4"].PutValue(430);

            Worksheet sheet2 = workbook.Worksheets.Add("Sales_Q2");
            sheet2.Cells["A1"].PutValue("Product");
            sheet2.Cells["B1"].PutValue("Amount");
            sheet2.Cells["A2"].PutValue("Apple");
            sheet2.Cells["B2"].PutValue(1500);
            sheet2.Cells["A3"].PutValue("Banana");
            sheet2.Cells["B3"].PutValue(950);
            sheet2.Cells["A4"].PutValue("Cherry");
            sheet2.Cells["B4"].PutValue(600);

            // Iterate through all worksheets and create a pie chart for each data table
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Determine the range of the data table.
                // For simplicity, assume data starts at A1, categories in column A, values in column B,
                // and the table ends at the last used row in column B.
                int lastRow = ws.Cells.MaxDataRow; // last row with data
                if (lastRow < 1) continue; // no data

                // Define the data range for values (e.g., B2:B{lastRow})
                string valueRange = $"B2:B{lastRow + 1}";
                // Define the category range (e.g., A2:A{lastRow})
                string categoryRange = $"A2:A{lastRow + 1}";

                // Add a pie chart to the worksheet.
                // Position the chart below the data table (adjust rows/columns as needed).
                int chartTopRow = lastRow + 3;
                int chartLeftColumn = 0;
                int chartBottomRow = chartTopRow + 15;
                int chartRightColumn = 7;

                int chartIndex = ws.Charts.Add(ChartType.Pie, chartTopRow, chartLeftColumn, chartBottomRow, chartRightColumn);
                Chart pieChart = ws.Charts[chartIndex];

                // Set the data source for the chart.
                pieChart.NSeries.Add(valueRange, true);
                pieChart.NSeries.CategoryData = categoryRange;

                // Optional: set chart title
                pieChart.Title.Text = $"Pie Chart - {ws.Name}";
                pieChart.ShowLegend = true;
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("BatchPieCharts.xlsx");
        }
    }
}
