// Title: C# – Update PivotTable Source Range Using Aspose.Cells ChangeDataSource
// Description: Demonstrates how to modify the source data of an existing PivotTable in a .NET workbook. The example creates initial data (A1:B4), builds a PivotTable, adds a new data block (C1:D4), calls PivotTable.ChangeDataSource with the new range and worksheet name, then refreshes and recalculates the pivot before saving the file as UpdatedPivotSource.xlsx.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | ChangeDataSource | update pivot source range | refresh pivot data | calculate pivot | programmatic pivot table | Excel automation | Aspose.Cells example | GitHub sample | code snippet
// Common Searches: Aspose.Cells change pivot table source range C# | How to use PivotTable.ChangeDataSource in .NET | Refresh PivotTable after source update Aspose.Cells | Update existing PivotTable data range programmatically | Aspose.Cells PivotTable example GitHub
// Developer Intent: Replace the data range of an existing PivotTable with a new worksheet range and refresh the pivot to reflect the updated data.
// Use Cases: Switch a PivotTable from its original range (A1:B4) to a new range (C1:D4) without recreating the pivot. | Programmatically adjust a PivotTable after inserting or modifying worksheet data. | Automate Excel reports where the source dataset changes dynamically and the pivot must stay in sync.
// AI Prompts: Show me C# code to change a PivotTable's source range with Aspose.Cells and refresh it. | Provide an Aspose.Cells example that updates a PivotTable to use a different worksheet range and saves the workbook. | Explain step‑by‑step how to modify the SourceData of an existing PivotTable in .NET and ensure the pivot reflects the new data.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to modify the source data of an existing PivotTable in a .NET workbook. The example creates initial data (A1:B4), builds a PivotTable, adds a new data block (C1:D4), calls PivotTable.ChangeDataSource with the new range and worksheet name, then refreshes and recalculates the pivot before saving the file as UpdatedPivotSource.xlsx.
class UpdatePivotSourceDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ---------- Original data (A1:B4) ----------
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1000);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(2000);
        sheet.Cells["A4"].PutValue("Orange");
        sheet.Cells["B4"].PutValue(3000);

        // Create a pivot table based on the original data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // ---------- New data (C1:D4) ----------
        sheet.Cells["C1"].PutValue("Product");
        sheet.Cells["D1"].PutValue("Sales");
        sheet.Cells["C2"].PutValue("Grape");
        sheet.Cells["D2"].PutValue(4000);
        sheet.Cells["C3"].PutValue("Mango");
        sheet.Cells["D3"].PutValue(5000);
        sheet.Cells["C4"].PutValue("Peach");
        sheet.Cells["D4"].PutValue(6000);

        // Change the pivot table's data source to the new range
        // The source array contains the range address and the worksheet name
        string[] newSource = new string[] { "C1:D4", sheet.Name };
        pivot.ChangeDataSource(newSource);

        // Refresh and recalculate the pivot table to reflect the new source
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook with the updated pivot table
        workbook.Save("UpdatedPivotSource.xlsx");
    }
}
