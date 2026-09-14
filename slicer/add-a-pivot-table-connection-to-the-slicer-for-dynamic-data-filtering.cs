// Title: Connect a slicer to a pivot table for interactive filtering using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, builds a pivot table from a data range, adds a slicer for a chosen field, and links the slicer to the pivot table. | Generate an Aspose.Cells example that adds multiple slicers for different pivot fields and connects each slicer to the same pivot table. | Provide C# code to save the workbook as an .xlsb file while preserving the slicer‑pivot relationship.
// Common Searches: aspocells c# connect slicer to pivot table for interactive filtering | example of adding slicer to Excel pivot table using Aspose.Cells .NET | how to programmatically bind a slicer to a pivot table in a .NET workbook
// Tags: Aspose.Cells add slicer to pivot table | C# pivot table slicer binding | Excel dynamic filtering with slicer using Aspose | save workbook as xlsb preserving slicer connection | programmatic slicer integration in .NET Excel

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a new workbook, populates a small data range, adds a pivot table, inserts a slicer for the Category field, connects the slicer to the pivot table for interactive filtering, and saves the file as PivotSlicerConnection.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate worksheet with sample data
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Product";
        cells["C1"].Value = "Sales";

        cells["A2"].Value = "Electronics";
        cells["B2"].Value = "Laptop";
        cells["C2"].Value = 1200;

        cells["A3"].Value = "Electronics";
        cells["B3"].Value = "Phone";
        cells["C3"].Value = 800;

        cells["A4"].Value = "Furniture";
        cells["B4"].Value = "Chair";
        cells["C4"].Value = 150;

        // Add a pivot table based on the data range
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:C4", "E2", "SalesPivot");
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the "Category" field of the pivot table
        int slicerIndex = sheet.Slicers.Add(pivot, "G2", "Category");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Connect the slicer to the pivot table for dynamic filtering
        slicer.AddPivotConnection(pivot);

        // Save the workbook
        workbook.Save("PivotSlicerConnection.xlsx");
    }
}
