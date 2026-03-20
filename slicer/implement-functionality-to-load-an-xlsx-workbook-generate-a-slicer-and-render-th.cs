using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Assume the data range A1:B5 contains the source data.
        // Create a table (ListObject) from this range.
        int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Add a slicer for the first column of the table.
        // The slicer will be placed with its upper‑left corner at cell D1.
        int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Optional: customize slicer appearance.
        slicer.Caption = "Category Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Save the workbook; the slicer is now rendered in the document.
        workbook.Save("output.xlsx");
    }
}