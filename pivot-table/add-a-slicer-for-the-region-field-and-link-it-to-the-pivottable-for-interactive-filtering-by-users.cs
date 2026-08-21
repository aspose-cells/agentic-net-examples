// Title: Add a Region slicer to a PivotTable using Aspose.Cells for .NET (C#)
// Description: Creates a workbook with product, region, and sales data, builds a PivotTable, adds a slicer for the "Region" field at G1, links the slicer to the PivotTable, refreshes it, and saves the file as PivotTableWithRegionSlicer.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel slicer | PivotTable slicer | region slicer | interactive filter | add slicer programmatically | link slicer to pivot | Aspose.Cells PivotTable example
// Common Searches: how to add a slicer to a PivotTable with Aspose.Cells C# | link slicer to PivotTable Aspose.Cells .NET | Aspose.Cells region slicer example | programmatic Excel slicer using Aspose.Cells | C# code to create interactive PivotTable filter
// Developer Intent: Generate a Region slicer, connect it to an existing PivotTable, and enable end‑user filtering in the resulting Excel workbook.
// Use Cases: Provide end‑users a clickable Region filter in generated reports. | Build dynamic dashboards where slicer selections instantly refresh PivotTable data. | Create reusable workbook templates with pre‑configured slicers for consistent reporting.
// AI Prompts: Show how to add multiple slicers for different fields to the same PivotTable with Aspose.Cells for .NET. | Generate code that changes a slicer's position, size, and style after workbook creation. | Explain how to synchronize slicer selections across several PivotTables on one worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Creates a workbook with product, region, and sales data, builds a PivotTable, adds a slicer for the "Region" field at G1, links the slicer to the PivotTable, refreshes it, and saves the file as PivotTableWithRegionSlicer.xlsx.
class AddRegionSlicerDemo
{
    // Entry point required for compilation
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that includes a "Region" field
        sheet.Cells["A1"].Value = "Product";
        sheet.Cells["B1"].Value = "Region";
        sheet.Cells["C1"].Value = "Sales";

        sheet.Cells["A2"].Value = "Apple";
        sheet.Cells["B2"].Value = "North";
        sheet.Cells["C2"].Value = 1200;

        sheet.Cells["A3"].Value = "Banana";
        sheet.Cells["B3"].Value = "South";
        sheet.Cells["C3"].Value = 800;

        sheet.Cells["A4"].Value = "Apple";
        sheet.Cells["B4"].Value = "South";
        sheet.Cells["C4"].Value = 500;

        // Create a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table: Product as row, Sales as data
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add a slicer for the "Region" field; place it starting at cell G1
        int slicerIndex = sheet.Slicers.Add(pivot, "G1", "Region");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Link the slicer to the pivot table and refresh
        slicer.AddPivotConnection(pivot);
        slicer.Refresh();

        // Save the workbook
        workbook.Save("PivotTableWithRegionSlicer.xlsx");
    }
}
