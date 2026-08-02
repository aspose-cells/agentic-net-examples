using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

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
        cells["A6"].Value = "North";
        cells["B6"].Value = 1300;
        cells["A7"].Value = "South";
        cells["B7"].Value = 1050;

        // Add a pivot table based on the data range
        int pivotIdx = sheet.PivotTables.Add("A1:B7", "D3", "SalesPivot");
        PivotTable pivot = sheet.PivotTables[pivotIdx];

        // Configure the pivot: Region as row field, Sales as data field
        pivot.AddFieldToArea(PivotFieldType.Row, "Region");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the Region field of the pivot table
        // Destination cell for the slicer's upper‑left corner is E3
        int slicerIdx = sheet.Slicers.Add(pivot, "E3", "Region");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Optional: customize slicer appearance
        slicer.Caption = "Region Filter";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.NumberOfColumns = 1;

        // Save the workbook
        workbook.Save("PivotWithRegionSlicer.xlsx");
    }
}