// Title: Create and link a Region slicer to a PivotTable in Aspose.Cells for C#
// AI Prompts: Write C# code with Aspose.Cells that adds a slicer for the 'Region' field, connects it to an existing PivotTable, refreshes the slicer, and saves the workbook. | Demonstrate how to programmatically attach a slicer to a PivotTable in Aspose.Cells .NET, set its position, refresh it after pivot calculations, and export the file.
// Common Searches: Aspose.Cells C# add slicer to pivot table example | How to link a slicer with a PivotTable using Aspose.Cells .NET | Refresh slicer after pivot calculation Aspose.Cells C# | Set slicer location and connect to pivot in Aspose.Cells workbook | Create region slicer for pivot table with Aspose.Cells API
// Tags: Aspose.Cells pivot table slicer | C# add slicer to pivot table | link slicer to pivot in Aspose.Cells | refresh slicer after pivot calculation | region field slicer Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerExample
{
    // Shows how to build a workbook, generate a PivotTable from sample data, add a slicer for the Region field, explicitly link the slicer to the PivotTable, refresh it, and save the result as PivotTableWithRegionSlicer.xlsx using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Product, Region, Sales
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1200;

            sheet.Cells["A3"].Value = "Apple";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 800;

            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = "North";
            sheet.Cells["C4"].Value = 600;

            sheet.Cells["A5"].Value = "Banana";
            sheet.Cells["B5"].Value = "South";
            sheet.Cells["C5"].Value = 900;

            // Add a PivotTable based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the PivotTable: Product as rows, Region as columns, Sales as data
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Region");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the PivotTable so that slicer can work with up‑to‑date data
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer for the "Region" field; place it starting at cell A1
            int slicerIndex = sheet.Slicers.Add(pivot, "A1", "Region");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Ensure the slicer is linked to the PivotTable (optional, but explicit)
            slicer.AddPivotConnection(pivot);

            // Refresh the slicer to reflect the current PivotTable state
            slicer.Refresh();

            // Save the workbook
            workbook.Save("PivotTableWithRegionSlicer.xlsx");
        }
    }
}
