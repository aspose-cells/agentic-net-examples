using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class UpdateSlicerStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(200);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount field
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Add a slicer linked to the Category field of the pivot table
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "F1", "Category");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Compute total amount from the source data
        double totalAmount = 0;
        for (int row = 1; row <= 3; row++)
        {
            totalAmount += worksheet.Cells[row, 1].DoubleValue;
        }

        // Update slicer style based on the total amount
        if (totalAmount > 300)
        {
            // Use a dark style when the total exceeds the threshold
            slicer.StyleType = SlicerStyleType.SlicerStyleDark2;
        }
        else
        {
            // Use a light style otherwise
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        }

        // Refresh the slicer to apply any pending changes
        slicer.Refresh();

        // Save the workbook
        workbook.Save("UpdatedSlicerStyle.xlsx");
    }
}