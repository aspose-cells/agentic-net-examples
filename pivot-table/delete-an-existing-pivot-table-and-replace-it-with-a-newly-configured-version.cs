// Title: Delete an existing PivotTable and create a new configured PivotTable in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Remove a specific PivotTable from a worksheet, then add a new PivotTable with a defined source range and set row and data fields using Aspose.Cells in C#. | Replace a PivotTable in an Excel file, configure its fields, refresh and calculate the data, and save the workbook programmatically with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells delete pivot table and add new one | Aspose.Cells replace pivot table in existing workbook | how to refresh and calculate a newly added PivotTable using Aspose.Cells .NET | programmatically remove first PivotTable and create another at D5 with Aspose.Cells
// Tags: remove pivot table Aspose.Cells C# | add pivot table Aspose.Cells source range | configure pivot fields Aspose.Cells .NET | refresh calculate pivot Aspose.Cells | replace pivot table workbook Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads input.xlsx, deletes the first PivotTable on the first worksheet, adds a new PivotTable at D5 using source range A1:B10, assigns 'Category' as a row field and 'Amount' as a data field, refreshes and calculates the pivot, then saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        // Load the workbook that contains the original PivotTable
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Access the collection of PivotTables on the worksheet
        PivotTableCollection pivotTables = sheet.PivotTables;

        // If there is at least one PivotTable, remove the first one (or any you need)
        if (pivotTables.Count > 0)
        {
            // Remove the existing PivotTable and its data
            PivotTable existingPivot = pivotTables[0];
            pivotTables.Remove(existingPivot);
        }

        // Add a new PivotTable
        // Source data range (adjust as needed)
        string sourceData = "A1:B10";
        // Destination cell for the top‑left corner of the new PivotTable
        string destCell = "D5";
        // Name of the new PivotTable
        string tableName = "NewPivotTable";

        int newPivotIndex = pivotTables.Add(sourceData, destCell, tableName);
        PivotTable newPivot = pivotTables[newPivotIndex];

        // Configure the new PivotTable (adjust field names to match your source data)
        newPivot.AddFieldToArea(PivotFieldType.Row, "Category");
        newPivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh and calculate the PivotTable to populate it with data
        newPivot.RefreshData();
        newPivot.CalculateData();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
