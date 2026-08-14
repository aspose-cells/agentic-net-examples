// Title: Add a Pivot Table Slicer with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates sample data, builds a pivot table, and inserts a slicer linked to the "Fruit" field. The slicer is placed at cell E2, its caption and style are customized, and the file is saved as an XLSX workbook.
// Keywords: Aspose.Cells slicer | C# pivot table slicer | Excel slicer programmatically | Aspose.Cells .NET example | interactive Excel filter | slicer style customization
// Common Searches: Aspose.Cells add slicer to pivot table C# | how to create slicer for pivot field using Aspose.Cells | set slicer caption and style Aspose.Cells | programmatic Excel slicer with Aspose | C# generate workbook with pivot slicer
// Developer Intent: Insert a slicer that is linked to a pivot table to allow users to filter data interactively.
// Use Cases: Provide end‑users a clickable filter for drill‑down analysis in generated reports. | Apply corporate branding to slicer captions and styles in automated dashboards. | Pre‑configure workbooks with slicers so no manual Excel setup is required.
// AI Prompts: Show C# code that adds a slicer to a pivot table using Aspose.Cells and sets its caption and style. | Generate an example that creates a pivot table from a range, attaches a slicer to a field, and saves the workbook. | Explain how to refresh a slicer after the underlying pivot data changes with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Creates a workbook, populates sample data, builds a pivot table, and inserts a slicer linked to the "Fruit" field. The slicer is placed at cell E2, its caption and style are customized, and the file is saved as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Fruit";
        cells["B1"].Value = "Quantity";
        cells["A2"].Value = "Apple";
        cells["B2"].Value = 10;
        cells["A3"].Value = "Orange";
        cells["B3"].Value = 5;
        cells["A4"].Value = "Banana";
        cells["B4"].Value = 8;

        // Add a pivot table that uses the data range A1:B4 and place it at D1
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: Fruit as row field, Quantity as data field
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
        pivot.RefreshData();
        pivot.CalculateData();

        // Insert a slicer linked to the pivot table for the "Fruit" field.
        // The slicer's upper‑left corner will start at cell E2.
        int slicerIndex = sheet.Slicers.Add(pivot, "E2", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Optional: customize slicer appearance
        slicer.Caption = "Fruit Filter";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Save the workbook to a file
        workbook.Save("SlicerWithPivot.xlsx", SaveFormat.Xlsx);
    }
}
