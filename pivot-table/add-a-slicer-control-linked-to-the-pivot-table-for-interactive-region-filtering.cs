// Title: Create an interactive Region slicer linked to a PivotTable using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, builds a PivotTable from Region and Sales data, and adds a slicer bound to the Region field with Aspose.Cells. | Demonstrate how to set the slicer style and caption after linking it to a PivotTable in Aspose.Cells. | Show the complete example that saves the workbook containing the PivotTable and its slicer to an .xlsx file.
// Common Searches: how to add a region slicer to a pivot table using Aspose.Cells C# | Aspose.Cells example linking slicer to pivot table for interactive filtering | customize slicer style and caption in Aspose.Cells .NET workbook
// Tags: Aspose.Cells add slicer to PivotTable | C# region field slicer Excel | pivot table interactive filter slicer .NET | custom slicer style Aspose.Cells | save workbook with slicer Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// // Generates a workbook, fills it with Region and Sales data, creates a PivotTable, inserts a slicer linked to the Region field, customizes its style and caption, and saves the file as PivotWithRegionSlicer.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data with a Region field
        cells["A1"].Value = "Region";
        cells["B1"].Value = "Sales";
        cells["A2"].Value = "North";
        cells["B2"].Value = 1200;
        cells["A3"].Value = "South";
        cells["B3"].Value = 950;
        cells["A4"].Value = "East";
        cells["B4"].Value = 800;
        cells["A5"].Value = "West";
        cells["B5"].Value = 1100;

        // Add a pivot table using the data range
        int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "RegionPivot");
        PivotTable pivot = sheet.PivotTables[pivotIdx];

        // Configure the pivot table: Region as row field, Sales as data field
        pivot.AddFieldToArea(PivotFieldType.Row, "Region");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the pivot table for interactive Region filtering
        // Using the overload Add(PivotTable, string destCellName, string baseFieldName)
        int slicerIdx = sheet.Slicers.Add(pivot, "F3", "Region");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Optional: customize slicer appearance
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.Caption = "Filter by Region";

        // Save the workbook
        workbook.Save("PivotWithRegionSlicer.xlsx");
    }
}
