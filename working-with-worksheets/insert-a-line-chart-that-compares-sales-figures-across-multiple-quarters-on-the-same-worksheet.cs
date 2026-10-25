// Title: Create a multi‑series line chart for quarterly product sales in an Aspose.Cells worksheet (C#)
// Common Searches: how to create a line chart with multiple series in Aspose.Cells using C# | Aspose.Cells C# example adding quarterly sales data to a line chart | setting category axis labels from a cell range in an Aspose.Cells line chart | populate worksheet with sales figures and generate a line chart programmatically using Aspose.Cells | Aspose.Cells NSeries.Add using cell range reference C#
// Tags: Aspose.Cells add line chart programmatically | C# define chart series from cell range | Aspose.Cells set X axis categories from worksheet | populate Excel worksheet with quarterly sales data | save workbook as Excel file using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, writes quarter labels and sales numbers for three products into cells A1:D5, adds a Line chart covering rows 6‑20, sets the chart title, adds three series referencing the product columns via cell ranges, assigns the quarter column as the X‑axis category labels, customizes each series name, and saves the file as 'QuarterlySalesLineChart.xlsx'.
class LineChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample sales data
            // Header row
            cells["A1"].PutValue("Quarter");
            cells["B1"].PutValue("Product A");
            cells["C1"].PutValue("Product B");
            cells["D1"].PutValue("Product C");

            // Data rows
            string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
            double[,] sales = {
                { 12000, 15000, 13000 },
                { 14000, 16000, 15500 },
                { 13000, 17000, 16500 },
                { 15000, 18000, 17500 }
            };

            for (int i = 0; i < quarters.Length; i++)
            {
                cells[i + 1, 0].PutValue(quarters[i]);               // Column A: Quarter
                cells[i + 1, 1].PutValue(sales[i, 0]);               // Column B: Product A
                cells[i + 1, 2].PutValue(sales[i, 1]);               // Column C: Product B
                cells[i + 1, 3].PutValue(sales[i, 2]);               // Column D: Product C
            }

            // Add a line chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Line, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Quarterly Sales Comparison";

            // Define the data range for categories (quarters)
            int firstDataRow = 1; // Row index of first data row (excluding header)
            int lastDataRow = firstDataRow + quarters.Length - 1;
            int categoryColumn = 0; // Column A

            // Add series for each product using the range address (RefersTo)
            chart.NSeries.Add(sheet.Cells.CreateRange(firstDataRow, 1, quarters.Length, 1).RefersTo, true); // Product A
            chart.NSeries.Add(sheet.Cells.CreateRange(firstDataRow, 2, quarters.Length, 1).RefersTo, true); // Product B
            chart.NSeries.Add(sheet.Cells.CreateRange(firstDataRow, 3, quarters.Length, 1).RefersTo, true); // Product C

            // Set the category (X) axis labels to the quarters
            chart.NSeries.CategoryData = sheet.Cells.CreateRange(firstDataRow, categoryColumn, quarters.Length, 1).RefersTo;

            // Optional: Customize series names
            chart.NSeries[0].Name = "Product A";
            chart.NSeries[1].Name = "Product B";
            chart.NSeries[2].Name = "Product C";

            // Save the workbook to a file
            workbook.Save("QuarterlySalesLineChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
