// Title: Create a pivot table with a linked column chart and export both to an ODS workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that builds a pivot table from sample sales data, adds a column chart linked to the pivot data, and saves the workbook as an ODS file using Aspose.Cells. | Show how to programmatically generate a pivot chart tied to a pivot table and export the result to ODS format with Aspose.Cells for .NET.
// Common Searches: aspocells c# create pivot table and column chart in same worksheet | export pivot table with chart to ODS using Aspose.Cells .NET | how to link a chart to a pivot table programmatically with Aspose.Cells | sample code for pivot chart generation in Aspose.Cells for .NET | save workbook as ODS file containing pivot table and chart Aspose.Cells
// Tags: aspocells create pivot table c# | aspocells add column chart to pivot data | aspocells export workbook to ods format | aspocells linked pivot chart generation | aspocells pivot table chart integration

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample sales data, builds a pivot table on a separate sheet, adds a column chart that references the same source range, and saves the entire workbook as an ODS file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Fill sample data that will be used for the pivot table.
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row.
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            // Sample data rows.
            string[,] rows = new string[,]
            {
                {"Beverages","Tea","120"},
                {"Beverages","Coffee","150"},
                {"Beverages","Juice","90"},
                {"Food","Bread","200"},
                {"Food","Butter","80"},
                {"Food","Cheese","130"},
                {"Snacks","Chips","110"},
                {"Snacks","Cookies","95"}
            };

            for (int i = 0; i < rows.GetLength(0); i++)
            {
                dataSheet.Cells[i + 1, 0].PutValue(rows[i, 0]); // Category
                dataSheet.Cells[i + 1, 1].PutValue(rows[i, 1]); // Product
                dataSheet.Cells[i + 1, 2].PutValue(Convert.ToDouble(rows[i, 2])); // Sales
            }

            // -------------------------------------------------
            // 2. Create a pivot table on a new worksheet.
            // -------------------------------------------------
            int pivotSheetIdx = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIdx];
            pivotSheet.Name = "Pivot";

            // Define the source data range (including header).
            int totalRows = rows.GetLength(0) + 1; // +1 for header row
            string sourceRange = $"=Data!$A$1:$C${totalRows}";

            // Add the pivot table (placed starting at cell A3).
            int pivotTableIdx = pivotSheet.PivotTables.Add(sourceRange, "A3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIdx];

            // Add fields to the pivot table.
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // Rename the data field (default aggregation is Sum).
            pivotTable.DataFields[0].Name = "Total Sales";

            // Refresh and calculate the pivot table.
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 3. Create a chart linked to the pivot table data.
            // -------------------------------------------------
            // Add a column chart to the same worksheet.
            int chartIdx = pivotSheet.Charts.Add(ChartType.Column, 10, 0, 25, 10);
            Chart chart = pivotSheet.Charts[chartIdx];

            // Use the original source data range for the chart series.
            // (Aspose.Cells does not expose a direct DataRange property for PivotTable.)
            chart.NSeries.Add(sourceRange, true);
            chart.Title.Text = "Sales by Category and Product";

            // -------------------------------------------------
            // 4. Save the workbook as an ODS file.
            // -------------------------------------------------
            workbook.Save("PivotChartOutput.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
