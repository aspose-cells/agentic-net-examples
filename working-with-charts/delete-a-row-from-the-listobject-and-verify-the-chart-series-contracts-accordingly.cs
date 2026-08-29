// Title: Delete a row from a ListObject table and automatically shrink the linked column chart series using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to remove a specific data row from a ListObject and have the associated column chart update its series range automatically in C#. | Programmatically delete a worksheet table row with Cells.DeleteRow and verify that the chart's NSeries point count reflects the change using Aspose.Cells for .NET. | Show how deleting a row from a ListObject updates both the table's DataRange and the linked column chart series without manual refresh in C#.
// Common Searches: aspocells delete row from listobject table and update column chart series | c# remove table row and keep column chart data synchronized using Aspose.Cells | how does Cells.DeleteRow affect chart NSeries in an Aspose.Cells workbook | example of deleting a ListObject row and adjusting chart series in C# | c# aspocells delete worksheet table row and automatically shrink chart series
// Tags: listobject row deletion auto-updates chart series | aspocells table modification reflects in column chart | c# cells.deleterow listobject chart synchronization | column chart series range adjustment after table row removal | aspocells workbook consistency between table and chart

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

// The example creates a workbook, defines a ListObject table, adds a column chart linked to the table data, deletes a data row using Cells.DeleteRow (which automatically updates the table), and then confirms that both the table row count and the chart series point count have decreased before saving the file.
class DeleteRowFromListObjectDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data for a table (ListObject)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Series1");
            cells["C1"].PutValue("Series2");

            cells["A2"].PutValue("Q1"); cells["B2"].PutValue(100); cells["C2"].PutValue(150);
            cells["A3"].PutValue("Q2"); cells["B3"].PutValue(200); cells["C3"].PutValue(250);
            cells["A4"].PutValue("Q3"); cells["B4"].PutValue(300); cells["C4"].PutValue(350);

            // Create a ListObject (table) that covers the data range A1:C4
            int tableIdx = worksheet.ListObjects.Add(0, 0, 3, 2, true);
            ListObject table = worksheet.ListObjects[tableIdx];
            table.ShowHeaderRow = true;

            // Add a column chart that uses the table data
            int chartIdx = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIdx];

            // Set the series data ranges (B2:B4 and C2:C4) and category data (A2:A4)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Display initial point count of the first series
            Console.WriteLine($"Initial points in first series: {chart.NSeries[0].Points.Count}");

            // Delete the second data row (Excel row 3, which is index 2 in zero‑based indexing)
            // This row belongs to the ListObject; deleting it updates the table automatically
            cells.DeleteRow(2, true);

            // Verify that the table now has one fewer data row
            Console.WriteLine($"Table data rows after deletion: {table.DataRange.RowCount}");

            // Verify that the chart series point count reflects the deleted row
            Console.WriteLine($"Points in first series after deletion: {chart.NSeries[0].Points.Count}");

            // Save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeleteRowFromListObjectDemo.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
