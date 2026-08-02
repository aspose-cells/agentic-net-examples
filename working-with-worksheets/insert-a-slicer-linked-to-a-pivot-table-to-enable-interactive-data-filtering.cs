// Title: Add a slicer to a pivot table using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, fills a simple data set, builds a pivot table, and inserts a slicer linked to the "Fruit" field. The slicer’s caption and style are customized before saving the file as an XLSX workbook.
// Keywords: Aspose.Cells slicer | C# pivot table slicer | programmatic Excel slicer .NET | add slicer to pivot table Aspose | customize slicer style Aspose.Cells
// Common Searches: Aspose.Cells add slicer to pivot table C# | how to link slicer with pivot table using Aspose | set slicer caption and style in Aspose.Cells | create interactive Excel dashboard with slicer .NET | save workbook with slicer Aspose.Cells
// Developer Intent: Insert a slicer that filters a pivot table field in a generated Excel file.
// Use Cases: Automated sales reports where users can filter by product category via a slicer. | Interactive Excel dashboards that combine pivot tables with slicers for quick data exploration. | Standardized workbook generation with branded slicer appearance across multiple files.
// AI Prompts: Generate C# code with Aspose.Cells that adds a slicer for the "Region" field of an existing pivot table and applies a dark style. | Provide a method to attach multiple slicers to different pivot fields, each with a unique caption and position. | Explain how to programmatically resize and reposition a slicer after it has been added to a worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Creates a new workbook, fills a simple data set, builds a pivot table, and inserts a slicer linked to the "Fruit" field. The slicer’s caption and style are customized before saving the file as an XLSX workbook.
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

        // Add a pivot table using the data range A1:B4, place it at D1, and name it "PivotTable1"
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: Fruit as row field, Quantity as data field
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Refresh and calculate the pivot data
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the pivot table for the "Fruit" field.
        // The slicer will be placed with its upper‑left corner at cell E1.
        int slicerIndex = sheet.Slicers.Add(pivot, "E1", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Optional: customize slicer appearance
        slicer.Caption = "Fruit Filter";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Save the workbook
        workbook.Save("SlicerPivotDemo.xlsx", SaveFormat.Xlsx);
    }
}
