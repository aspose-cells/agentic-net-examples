// Title: Aspose.Cells for .NET – Add Multiple Series to a Column Chart and Freeze Source Rows
// Description: Shows how to build a workbook, fill rows 1‑6 with category labels and three numeric series, create a column chart with three vertical series, assign series names from the header row, freeze the first six rows using FreezePanes, and save the result as MultipleSeriesWithFreeze.xlsx.
// Keywords: Aspose.Cells | .NET | C# | column chart | multiple series | NSeries.Add | set series names | FreezePanes | freeze rows | Excel automation | chart data source
// Common Searches: Aspose.Cells add multiple series to chart | FreezePanes rows Aspose.Cells .NET | set chart series names from header Aspose.Cells | create column chart with categories Aspose.Cells | freeze top rows in Excel using Aspose.Cells
// Developer Intent: Create a column chart with several data series and keep the source rows visible by freezing them.
// Use Cases: Generate a month‑by‑month column chart for three products while the category and value rows stay fixed during scrolling. | Automatically pull series names from the header row so the chart legend updates when the header changes. | Freeze the first six rows of a worksheet to provide constant reference to source data in large spreadsheets.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart that pulls four series from columns B‑E and freezes rows 1‑5. | Explain each parameter of the FreezePanes method and how it determines the frozen area in Aspose.Cells for .NET. | Show how to iterate over all data columns in a worksheet and add each as a separate series to a chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultipleSeriesWithFreeze
{
    // Shows how to build a workbook, fill rows 1‑6 with category labels and three numeric series, create a column chart with three vertical series, assign series names from the header row, freeze the first six rows using FreezePanes, and save the result as MultipleSeriesWithFreeze.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data
                // -------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["D1"].PutValue("Series 3");

                // Data rows (rows 2 to 6)
                string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
                int[,] values = {
                    { 10, 15, 20 },
                    { 20, 25, 30 },
                    { 30, 35, 40 },
                    { 40, 45, 50 },
                    { 50, 55, 60 }
                };

                for (int i = 0; i < categories.Length; i++)
                {
                    int row = i + 2; // Data starts at row 2
                    sheet.Cells[$"A{row}"].PutValue(categories[i]);
                    sheet.Cells[$"B{row}"].PutValue(values[i, 0]);
                    sheet.Cells[$"C{row}"].PutValue(values[i, 1]);
                    sheet.Cells[$"D{row}"].PutValue(values[i, 2]);
                }

                // -------------------------------------------------
                // Add a chart
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 22, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the category (X‑axis) data for all series
                chart.NSeries.CategoryData = "A2:A6";

                // Add multiple series (B, C, D columns) using the SeriesCollection.Add method
                // Each call adds one series; the range is vertical (isVertical = true)
                chart.NSeries.Add("B2:B6", true); // Series 1
                chart.NSeries.Add("C2:C6", true); // Series 2
                chart.NSeries.Add("D2:D6", true); // Series 3

                // Optionally set series names from the header row
                chart.NSeries.SetSeriesNames(0, "B1:D1", true);

                // -------------------------------------------------
                // Freeze the rows that contain the series data (rows 1‑6)
                // FreezePanes freezes rows above the specified row index.
                // Use the 4‑parameter overload for compatibility with all Aspose.Cells versions.
                // Parameters: row, column, totalRows, totalColumns.
                // Setting totalRows and totalColumns to 0 keeps the default scrolling area.
                // -------------------------------------------------
                sheet.FreezePanes(7, 0, 0, 0);

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("MultipleSeriesWithFreeze.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
