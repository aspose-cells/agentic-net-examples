// Title: Move a pivot column field to the row area in an Aspose.Cells workbook with C#
// AI Prompts: Write C# code using Aspose.Cells to transfer a pivot table field from the column area to the row area. | Show how to programmatically reassign a pivot field's parent area from Column to Row in Aspose.Cells for .NET. | Demonstrate removing a column field and adding it as a row field in an Aspose.Cells pivot table.
// Common Searches: Aspose.Cells C# change pivot table field from column to row | programmatically move pivot column field to rows using Aspose.Cells .NET | how to relocate a pivot field area in Aspose.Cells workbook | C# Aspose.Cells pivot table rearrange fields example | remove and add pivot field to different area Aspose.Cells
// Tags: Aspose.Cells pivot field area change | C# pivot column to row conversion | Aspose.Cells rearrange pivot fields | programmatic pivot layout update .NET | remove and add pivot field Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, adds sample data, builds a pivot table with Product rows, Region columns, and Sales data, then moves the Region field from the column area to the row area by removing and re‑adding it, refreshes the pivot, and saves the workbook as PivotItemMoved.xlsx.
class MovePivotItemDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Region");
        sheet.Cells["C1"].PutValue("Sales");

        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue("North");
        sheet.Cells["C2"].PutValue(100);

        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue("South");
        sheet.Cells["C3"].PutValue(150);

        sheet.Cells["A4"].PutValue("Apple");
        sheet.Cells["B4"].PutValue("South");
        sheet.Cells["C4"].PutValue(200);

        sheet.Cells["A5"].PutValue("Banana");
        sheet.Cells["B5"].PutValue("North");
        sheet.Cells["C5"].PutValue(250);

        // Add a pivot table based on the data range
        int ptIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
        PivotTable pt = sheet.PivotTables[ptIndex];

        // Initial layout: Product in rows, Region in columns, Sales as data
        pt.AddFieldToArea(PivotFieldType.Row, "Product");
        pt.AddFieldToArea(PivotFieldType.Column, "Region");
        pt.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Populate the pivot table
        pt.RefreshData();
        pt.CalculateData();

        // ----- Move the "Region" field from Column area to Row area -----
        // Remove the field from the column area
        pt.RemoveField(PivotFieldType.Column, "Region");
        // Add the same field to the row area
        pt.AddFieldToArea(PivotFieldType.Row, "Region");

        // Recalculate after moving the field
        pt.RefreshData();
        pt.CalculateData();

        // Save the workbook with the updated pivot layout
        wb.Save("PivotItemMoved.xlsx");
    }
}
