// Title: How to programmatically add a slicer to every pivot table in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that iterates through all worksheets in a Workbook and adds a slicer for the first base field of each PivotTable using Aspose.Cells, positioning each slicer in column A with a 5‑row offset. | Demonstrate how to set a custom caption and apply a light style to each slicer created in a loop over pivot tables with Aspose.Cells.
// Common Searches: C# Aspose.Cells add slicer to each pivot table in a workbook | loop through worksheets and create slicers for pivot tables using Aspose.Cells .NET | position slicer in specific cell with row offset when generating Excel file programmatically | set slicer style and caption in Aspose.Cells C# example
// Tags: Aspose.Cells add slicer to pivot table | C# loop worksheets create slicers | Excel slicer placement column A Aspose.Cells | slicer style caption Aspose.Cells | batch slicer generation Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerBatchDemo
{
    // The example creates a workbook, adds sample data and a pivot table, then loops through every worksheet and each pivot table to insert a slicer for the first base field. Each slicer is placed in column A with a 5‑row gap, given a custom caption, styled with a light theme, and the workbook is saved as BatchSlicersDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example data for demonstration – create a worksheet with a pivot table
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";
            dataSheet.Cells["A1"].Value = "Category";
            dataSheet.Cells["B1"].Value = "Value";
            dataSheet.Cells["A2"].Value = "A";
            dataSheet.Cells["B2"].Value = 10;
            dataSheet.Cells["A3"].Value = "B";
            dataSheet.Cells["B3"].Value = 20;
            dataSheet.Cells["A4"].Value = "A";
            dataSheet.Cells["B4"].Value = 30;

            // Add a pivot table to the same sheet (for demo purposes)
            int pivotIdx = dataSheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = dataSheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Loop through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Loop through each pivot table in the current worksheet
                foreach (PivotTable pt in sheet.PivotTables)
                {
                    // Ensure the pivot table has at least one base field to use for the slicer
                    if (pt.BaseFields.Count == 0)
                        continue;

                    // Use the first base field name as the slicer field
                    string baseFieldName = pt.BaseFields[0].Name;

                    // Determine a destination cell for the slicer.
                    // Here we place each slicer starting from column A and offset rows to avoid overlap.
                    // The offset is based on the current count of slicers already on the sheet.
                    int slicerCount = sheet.Slicers.Count;
                    int startRow = slicerCount * 5; // 5 rows gap between slicers
                    string destCell = CellsHelper.CellIndexToName(startRow, 0); // column A

                    // Add the slicer using the (PivotTable, destCellName, baseFieldName) overload
                    int slicerIndex = sheet.Slicers.Add(pt, destCell, baseFieldName);
                    Slicer slicer = sheet.Slicers[slicerIndex];

                    // Optional: set a friendly caption and style
                    slicer.Caption = $"{pt.Name} - {baseFieldName}";
                    slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
                }
            }

            // Save the workbook with the created slicers
            workbook.Save("BatchSlicersDemo.xlsx");
        }
    }
}
