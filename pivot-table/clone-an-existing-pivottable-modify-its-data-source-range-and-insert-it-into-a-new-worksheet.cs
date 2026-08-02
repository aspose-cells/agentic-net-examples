// Title: Clone a PivotTable, modify its data source, and place it on a new worksheet – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook with sample data, build an original PivotTable, clone it to a newly added worksheet using the Add(PivotTable, string, string) overload, change the cloned table's data source with ChangeDataSource, refresh and recalculate the pivot, and save the file containing both tables.
// Keywords: Aspose.Cells | C# | .NET | PivotTable clone | ChangeDataSource | Add pivot to new sheet | RefreshData | CalculateData | Excel automation | workbook pivot example
// Common Searches: Aspose.Cells clone PivotTable C# | Change data source of a cloned pivot Aspose.Cells | Add PivotTable to another worksheet using Aspose.Cells | Refresh and recalculate cloned PivotTable .NET | Aspose.Cells ChangeDataSource method example
// Developer Intent: Copy an existing PivotTable, assign a different data range, and insert the copy into a separate worksheet programmatically.
// Use Cases: Create a summary sheet that shows a pivot based on a subset of rows from the original data. | Generate multiple pivot reports from the same workbook, each using a distinct data slice. | Automate report generation where each worksheet contains a cloned pivot reflecting a specific timeframe or category.
// AI Prompts: Write C# code with Aspose.Cells to clone a PivotTable, set a new data range, and add it to a new worksheet. | Explain how to refresh and recalculate a cloned PivotTable after changing its source using Aspose.Cells. | Show the syntax for Add(PivotTable, string, string) and ChangeDataSource together in an Aspose.Cells .NET example.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook with sample data, build an original PivotTable, clone it to a newly added worksheet using the Add(PivotTable, string, string) overload, change the cloned table's data source with ChangeDataSource, refresh and recalculate the pivot, and save the file containing both tables.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data on the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "SourceData";

        // Header row
        sourceSheet.Cells["A1"].PutValue("Category");
        sourceSheet.Cells["B1"].PutValue("Product");
        sourceSheet.Cells["C1"].PutValue("Sales");

        // Populate rows 2‑10 with sample data
        for (int i = 2; i <= 10; i++)
        {
            sourceSheet.Cells[$"A{i}"].PutValue("Cat" + (i % 3 + 1));
            sourceSheet.Cells[$"B{i}"].PutValue("Prod" + i);
            sourceSheet.Cells[$"C{i}"].PutValue(i * 100);
        }

        // Create the original pivot table on the source sheet
        PivotTableCollection sourcePivots = sourceSheet.PivotTables;
        int sourcePivotIdx = sourcePivots.Add("=SourceData!A1:C10", "E1", "OriginalPivot");
        PivotTable sourcePivot = sourcePivots[sourcePivotIdx];

        // Configure fields: Category & Product as rows, Sales as data
        sourcePivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category column
        sourcePivot.AddFieldToArea(PivotFieldType.Row, 1);   // Product column
        sourcePivot.AddFieldToArea(PivotFieldType.Data, 2);  // Sales column
        sourcePivot.CalculateData();

        // Add a new worksheet that will host the cloned pivot table
        Worksheet clonedSheet = workbook.Worksheets.Add("ClonedPivotSheet");

        // Clone the existing pivot table into the new sheet using Add(PivotTable, string, string)
        PivotTableCollection clonedPivots = clonedSheet.PivotTables;
        int clonedPivotIdx = clonedPivots.Add(sourcePivot, "A1", "ClonedPivot");
        PivotTable clonedPivot = clonedPivots[clonedPivotIdx];

        // Change the data source of the cloned pivot table to a different range (rows 1‑5)
        // The array format is { "RangeAddress", "SheetName" }
        string[] newDataSource = new string[] { "A1:C5", "SourceData" };
        clonedPivot.ChangeDataSource(newDataSource);

        // Refresh and recalculate the cloned pivot to reflect the new source
        clonedPivot.RefreshData();
        clonedPivot.CalculateData();

        // Save the workbook with both pivot tables
        workbook.Save("ClonedPivotExample.xlsx");
    }
}
