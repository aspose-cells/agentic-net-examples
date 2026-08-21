// Title: Hide Zero‑Value Rows in an Aspose.Cells Pivot Table (C#)
// Description: Creates a workbook with sample sales data, builds a pivot table, and sets the ShowEmptyRow property to false so rows whose aggregated value is zero are omitted before saving the file.
// Keywords: Aspose.Cells hide zero rows | ShowEmptyRow property | Aspose.Cells pivot table filter | C# remove empty rows pivot | Aspose.Cells .NET zero‑value rows
// Common Searches: Aspose.Cells hide zero rows in pivot | ShowEmptyRow example C# | remove empty rows from Aspose pivot table | filter zero values Aspose.Cells pivot | C# Aspose.Cells pivot table hide empty rows
// Developer Intent: Exclude rows with zero or empty aggregated values from a generated pivot table using Aspose.Cells.
// Use Cases: Generate a sales summary that automatically skips categories with no revenue. | Create a product performance report that displays only items with non‑zero sales. | Produce a clean financial pivot view by removing rows that total zero after refresh.
// AI Prompts: Write C# code with Aspose.Cells to hide rows that have zero totals in a pivot table. | Explain the effect of the ShowEmptyRow property on pivot table output in Aspose.Cells. | Provide a step‑by‑step example that filters out zero‑value rows after refreshing a pivot table using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotZeroRowHideDemo
{
    // Creates a workbook with sample sales data, builds a pivot table, and sets the ShowEmptyRow property to false so rows whose aggregated value is zero are omitted before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data with some zero‑value rows
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Electronics");
            dataSheet.Cells["B2"].PutValue("TV");
            dataSheet.Cells["C2"].PutValue(1500);

            dataSheet.Cells["A3"].PutValue("Electronics");
            dataSheet.Cells["B3"].PutValue("Radio");
            dataSheet.Cells["C3"].PutValue(0); // zero‑value row

            dataSheet.Cells["A4"].PutValue("Furniture");
            dataSheet.Cells["B4"].PutValue("Chair");
            dataSheet.Cells["C4"].PutValue(800);

            dataSheet.Cells["A5"].PutValue("Furniture");
            dataSheet.Cells["B5"].PutValue("Table");
            dataSheet.Cells["C5"].PutValue(0); // zero‑value row

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table based on the data range
            int pivotIndex = pivotSheet.PivotTables.Add(
                "=Sheet1!A1:C5", // source range
                "A1",            // destination cell
                "PivotTable1");  // pivot table name

            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, columns = Product, data = Sales (sum)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Category
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Sales

            // Hide rows that have no data (including zero‑value rows)
            // Setting ShowEmptyRow to false removes rows where the aggregated value is zero/empty.
            pivotTable.ShowEmptyRow = false;

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_HideZeroRows.xlsx");
        }
    }
}
