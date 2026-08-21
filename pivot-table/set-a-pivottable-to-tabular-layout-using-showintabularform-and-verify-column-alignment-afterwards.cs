// Title: Set PivotTable to Tabular Layout with ShowInTabularForm and Validate Column Alignment (Aspose.Cells for .NET)
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable on a separate sheet, switch the layout to Tabular form using ShowInTabularForm, refresh and calculate the pivot, then programmatically verify that the ColumnRange start column matches the DataBodyRange start column before saving the file.
// Keywords: Aspose.Cells | ShowInTabularForm | PivotTable Tabular layout | ColumnRange | DataBodyRange | C# .NET Excel automation | pivot column alignment verification
// Common Searches: Aspose.Cells set pivot table tabular form | ShowInTabularForm example C# | verify pivot column alignment Aspose.Cells | PivotTable ColumnRange vs DataBodyRange | C# code to apply tabular layout to Excel pivot
// Developer Intent: Apply Tabular layout to a PivotTable and programmatically confirm that the column area aligns with the data body area.
// Use Cases: Generate reports where a flat row hierarchy improves readability, requiring Tabular layout and automatic validation of column positions. | Create Excel workbooks for financial or inventory analysis that must maintain consistent column alignment after applying PivotTable formatting.
// AI Prompts: Write C# code using Aspose.Cells that calls ShowInTabularForm on a PivotTable, refreshes it, and throws an exception if ColumnRange.StartColumn differs from DataBodyRange.StartColumn. | Provide a reusable method that accepts a PivotTable object, sets the tabular layout, and returns a boolean indicating column‑alignment success. | Explain the effect of ShowInTabularForm on PivotTable structure and show how to verify alignment with ColumnRange and DataBodyRange properties in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add sample data, build a PivotTable on a separate sheet, switch the layout to Tabular form using ShowInTabularForm, refresh and calculate the pivot, then programmatically verify that the ColumnRange start column matches the DataBodyRange start column before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet for source data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate sample data for the pivot table
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Item");
        dataSheet.Cells["C1"].PutValue("Amount");

        dataSheet.Cells["A2"].PutValue("Fruit");
        dataSheet.Cells["B2"].PutValue("Apple");
        dataSheet.Cells["C2"].PutValue(10);

        dataSheet.Cells["A3"].PutValue("Fruit");
        dataSheet.Cells["B3"].PutValue("Banana");
        dataSheet.Cells["C3"].PutValue(20);

        dataSheet.Cells["A4"].PutValue("Vegetable");
        dataSheet.Cells["B4"].PutValue("Carrot");
        dataSheet.Cells["C4"].PutValue(15);

        // Add a new worksheet to host the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table (source range, destination cell, name)
        int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C4", "A1", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Add fields to the pivot table: two row fields and one data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Item");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Set the layout to Tabular form using the provided method
        pivotTable.ShowInTabularForm();

        // Refresh data from the source and calculate the pivot table values
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Verify column alignment after applying Tabular layout
        // In Tabular form the start column of the column area should match the start column of the data body area
        CellArea columnRange = pivotTable.ColumnRange;
        CellArea dataBodyRange = pivotTable.DataBodyRange;

        bool isAligned = columnRange.StartColumn == dataBodyRange.StartColumn;
        Console.WriteLine("Column alignment verification: " + (isAligned ? "Passed" : "Failed"));

        // Save the workbook to a file
        workbook.Save("PivotTableTabularFormDemo.xlsx");
    }
}
