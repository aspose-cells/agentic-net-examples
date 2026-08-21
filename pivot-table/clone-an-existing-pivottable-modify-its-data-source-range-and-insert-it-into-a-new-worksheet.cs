// Title: Clone a PivotTable, change its data source, and add it to a new sheet with Aspose.Cells for .NET
// Description: C# example that creates a workbook, builds an original PivotTable, clones it to a newly added worksheet, updates the cloned table's data source to a different range on the same sheet, refreshes and recalculates the pivot, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# PivotTable clone | ChangeDataSource Aspose.Cells | add pivot table to new worksheet .NET | duplicate pivot table C# | refresh cloned pivot Aspose | .NET Excel pivot example
// Common Searches: how to clone a pivot table with Aspose.Cells C# | change data source of a copied pivot table Aspose.Cells | add cloned pivot table to another sheet using Aspose.Cells for .NET | Aspose.Cells example for duplicating pivot tables | refresh pivot after changing source range Aspose.Cells
// Developer Intent: Copy an existing PivotTable, point it to a new data range, and place the copy on a separate worksheet.
// Use Cases: Create a summary sheet by reusing a template PivotTable with a different dataset. | Generate multiple comparative reports by cloning a pivot and assigning each clone a unique source range. | Automate workbook generation where a master PivotTable is duplicated across sheets for department‑specific analysis.
// AI Prompts: Write C# code with Aspose.Cells to clone a PivotTable, set a new data source range, and insert it into a new worksheet. | Explain the parameters required for the ChangeDataSource method when updating a cloned PivotTable in Aspose.Cells. | Show how to programmatically refresh and recalculate a cloned PivotTable after modifying its source range using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that creates a workbook, builds an original PivotTable, clones it to a newly added worksheet, updates the cloned table's data source to a different range on the same sheet, refreshes and recalculates the pivot, and saves the file using Aspose.Cells for .NET.
class ClonePivotTableExample
{
    static void Main()
    {
        // Create a workbook and add source data
        Workbook workbook = new Workbook();
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "SourceData";

        // Populate sample data for the original pivot table
        sourceSheet.Cells["A1"].PutValue("Category");
        sourceSheet.Cells["B1"].PutValue("Product");
        sourceSheet.Cells["C1"].PutValue("Sales");
        for (int i = 2; i <= 6; i++)
        {
            sourceSheet.Cells[$"A{i}"].PutValue("Cat" + (i % 3 + 1));
            sourceSheet.Cells[$"B{i}"].PutValue("Prod" + i);
            sourceSheet.Cells[$"C{i}"].PutValue(i * 100);
        }

        // Create the original pivot table on the source sheet
        PivotTableCollection sourcePivots = sourceSheet.PivotTables;
        int srcPivotIdx = sourcePivots.Add("=SourceData!A1:C6", "E1", "OriginalPivot");
        PivotTable srcPivot = sourcePivots[srcPivotIdx];
        srcPivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row
        srcPivot.AddFieldToArea(PivotFieldType.Data, 2);  // Sales as data
        srcPivot.CalculateData();

        // Add a new worksheet where the cloned pivot table will reside
        Worksheet clonedSheet = workbook.Worksheets.Add("ClonedPivotSheet");

        // Clone the existing pivot table into the new worksheet
        PivotTableCollection clonedPivots = clonedSheet.PivotTables;
        int clonedIdx = clonedPivots.Add(srcPivot, "A1", "ClonedPivot");
        PivotTable clonedPivot = clonedPivots[clonedIdx];

        // Change the data source of the cloned pivot table
        // The array contains the new range address and the sheet name
        string[] newSource = new string[] { "D1:F6", "SourceData" };
        clonedPivot.ChangeDataSource(newSource);

        // Refresh the cloned pivot table to apply the new data source
        clonedPivot.RefreshData();
        clonedPivot.CalculateData();

        // Save the workbook with both pivot tables
        workbook.Save("ClonedPivotExample.xlsx");
    }
}
