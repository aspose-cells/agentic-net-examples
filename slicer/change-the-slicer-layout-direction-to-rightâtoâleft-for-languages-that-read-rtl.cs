// Title: Set a slicer’s layout direction to right‑to‑left for RTL languages using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a pivot table, adds a slicer, and sets the slicer’s TextDirection property to RightToLeft with Aspose.Cells. | Show how to enable right‑to‑left display for a worksheet and its linked slicer in an Aspose.Cells workbook. | Provide an example that configures a slicer’s orientation to RTL, saves the workbook as an .xlsx file, and demonstrates optional worksheet RTL settings.
// Common Searches: Aspose.Cells C# set slicer text direction right to left | How to make Excel slicer RTL using Aspose.Cells .NET | Enable right‑to‑left layout for slicer in Aspose.Cells workbook | C# Aspose.Cells pivot table slicer orientation RTL | Display worksheet right to left with slicer in Aspose.Cells
// Tags: Aspose.Cells slicer right-to-left layout | C# set slicer TextDirection RTL | Aspose.Cells pivot table slicer orientation | display worksheet right-to-left Aspose.Cells | Excel slicer RTL Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the 'Category' field, sets the slicer’s TextDirection to RightToLeft, optionally enables right‑to‑left display for the worksheet, and saves the file as SlicerRTL.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue("Alpha");
        cells["B2"].PutValue(100);
        cells["A3"].PutValue("Beta");
        cells["B3"].PutValue(200);
        cells["A4"].PutValue("Gamma");
        cells["B4"].PutValue(300);

        // Add a pivot table based on the sample data
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Add a slicer linked to the pivot table for the "Category" field
        SlicerCollection slicers = worksheet.Slicers;
        int slicerIndex = slicers.Add(pivotTable, "E1", "Category");
        Slicer slicer = slicers[slicerIndex];

        // Change the slicer layout direction to right‑to‑left
        slicer.Shape.TextDirection = TextDirectionType.RightToLeft;

        // (Optional) Set the whole worksheet to display right‑to‑left
        worksheet.DisplayRightToLeft = true;

        // Save the workbook
        workbook.Save("SlicerRTL.xlsx");
    }
}
