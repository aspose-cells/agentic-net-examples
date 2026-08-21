// Title: Create a Pivot Table and Freeze Row‑Label Column with Aspose.Cells for .NET (C#)
// Description: This example builds a new workbook, adds a data sheet with sample sales records, creates a pivot table on that range, assigns Category and Product as row fields and Sales as a data field, refreshes and calculates the pivot, then freezes the column that holds the row labels so it stays visible while scrolling, and finally saves the file as an XLSX document.
// Keywords: Aspose.Cells pivot table C# | freeze row labels Aspose.Cells | FreezePanes pivot table .NET | Excel pivot freeze column example | Aspose.Cells create and freeze pivot | C# Aspose.Cells pivot freeze panes
// Common Searches: how to freeze row‑label column in a pivot table using Aspose.Cells | Aspose.Cells C# example for pivot table with frozen panes | freeze panes for pivot tables in .NET | Aspose.Cells create pivot and lock row labels | C# code to add pivot table and apply FreezePanes
// Developer Intent: Generate an Excel workbook that contains a pivot table whose row‑label column remains fixed during horizontal scrolling.
// Use Cases: Produce a sales summary where categories and products are row labels that stay in view while users explore large data sets. | Build an interactive dashboard that adds a pivot table to a separate worksheet and locks the label column for better navigation. | Export analytical results to Excel with a pivot table that keeps its row identifiers visible as the data area scrolls.
// AI Prompts: Write C# code with Aspose.Cells to create a pivot table from a data range and freeze the column that contains the row labels. | Show an Aspose.Cells example that adds a pivot table, sets row and data fields, refreshes the cache, and applies FreezePanes to keep row labels visible. | Explain how to calculate the correct column index for freezing the row‑label column in a pivot table created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotFreezeDemo
{
    // This example builds a new workbook, adds a data sheet with sample sales records, creates a pivot table on that range, assigns Category and Product as row fields and Sales as a data field, refreshes and calculates the pivot, then freezes the column that holds the row labels so it stays visible while scrolling, and finally saves the file as an XLSX document.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------- Create data worksheet ----------
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data for the pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Sales");

                dataSheet.Cells["A2"].PutValue("Electronics");
                dataSheet.Cells["B2"].PutValue("Laptop");
                dataSheet.Cells["C2"].PutValue(1200);

                dataSheet.Cells["A3"].PutValue("Electronics");
                dataSheet.Cells["B3"].PutValue("Phone");
                dataSheet.Cells["C3"].PutValue(800);

                dataSheet.Cells["A4"].PutValue("Furniture");
                dataSheet.Cells["B4"].PutValue("Chair");
                dataSheet.Cells["C4"].PutValue(150);

                dataSheet.Cells["A5"].PutValue("Furniture");
                dataSheet.Cells["B5"].PutValue("Table");
                dataSheet.Cells["C5"].PutValue(300);

                // ---------- Create pivot table worksheet ----------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Add a pivot table; source range is the whole data area, destination cell is A1 in the pivot sheet
                int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot cache and calculate the data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // ---------- Freeze the row‑label column ----------
                // Freeze the column that contains row labels.
                int freezeColumn = pivotTable.RowRange.StartColumn + 1; // column index after frozen area
                // FreezePanes(row, column, totalRows, totalColumns)
                pivotSheet.FreezePanes(0, freezeColumn, 0, 0);

                // Save the workbook
                workbook.Save("PivotTableWithFrozenRowLabels.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
