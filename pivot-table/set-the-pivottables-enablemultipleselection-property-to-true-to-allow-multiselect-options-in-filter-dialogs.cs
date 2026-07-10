using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Fruit";
        sheet.Cells["B1"].Value = "Quantity";
        sheet.Cells["A2"].Value = "Apple";
        sheet.Cells["B2"].Value = 10;
        sheet.Cells["A3"].Value = "Orange";
        sheet.Cells["B3"].Value = 15;
        sheet.Cells["A4"].Value = "Banana";
        sheet.Cells["B4"].Value = 20;

        // Add a pivot table to the worksheet
        int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Add fields to the pivot table (row and data)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Add a page field so that multiple selection can be enabled
        pivotTable.AddFieldToArea(PivotFieldType.Page, "Fruit");

        // Enable multiple item selection on the page field
        PivotField pageField = pivotTable.PageFields[0];
        pageField.IsMultipleItemSelectionAllowed = true;

        // Save the workbook
        workbook.Save("PivotTable_MultiSelect.xlsx");
    }
}