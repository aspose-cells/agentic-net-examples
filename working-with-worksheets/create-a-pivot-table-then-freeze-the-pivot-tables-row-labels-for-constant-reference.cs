// Title: Create a Pivot Table and Freeze Row Labels (first column) with Aspose.Cells for .NET (C#)
// Description: Shows how to generate a workbook, populate a data sheet, add a pivot table, place Category and Item in the row area, Amount in the data area, switch to tabular layout, refresh and calculate the pivot, then freeze the first column (row labels) on the pivot worksheet, and finally save the file as an .xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | pivot table | freeze panes | freeze row labels | first column freeze | tabular form pivot | programmatic Excel | Excel automation | add pivot table Aspose
// Common Searches: Aspose.Cells freeze first column in pivot table | C# create pivot table and lock row labels | How to freeze row labels of a pivot sheet using Aspose.Cells | Programmatically add pivot table and freeze panes .NET | Aspose.Cells tabular layout pivot example
// Developer Intent: Generate a pivot table from a data range and keep the row‑label column fixed while scrolling, using Aspose.Cells for .NET.
// Use Cases: Build a sales‑by‑category report where the Category column stays visible during horizontal scrolling. | Create a dynamic inventory dashboard with row headers locked for easier navigation. | Produce a printable Excel workbook that shows a tabular‑layout pivot table with the first column frozen for consistent reference.
// AI Prompts: Write C# code with Aspose.Cells to create a pivot table from a range and freeze the first column on the pivot sheet. | Show how to set a pivot table to tabular form, refresh its data, and apply FreezePanes to keep row labels visible. | Explain the steps to add row and data fields to a pivot table and lock the row‑label column using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to generate a workbook, populate a data sheet, add a pivot table, place Category and Item in the row area, Amount in the data area, switch to tabular layout, refresh and calculate the pivot, then freeze the first column (row labels) on the pivot worksheet, and finally save the file as an .xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the default worksheet for source data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Item");
            dataSheet.Cells["C1"].PutValue("Amount");

            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(120);

            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(80);

            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(60);

            dataSheet.Cells["A5"].PutValue("Vegetable");
            dataSheet.Cells["B5"].PutValue("Broccoli");
            dataSheet.Cells["C5"].PutValue(90);

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Add a pivot table: source range, destination cell, and name
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Item");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Layout the pivot table in tabular form
            pivotTable.ShowInTabularForm();

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Freeze the first column (column A) while keeping all rows scrollable
            // Parameters: row, column, totalRows, totalColumns
            pivotSheet.FreezePanes(0, 1, 0, 1);

            // Define output file path
            string outputPath = "PivotTableWithFrozenRowLabels.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
