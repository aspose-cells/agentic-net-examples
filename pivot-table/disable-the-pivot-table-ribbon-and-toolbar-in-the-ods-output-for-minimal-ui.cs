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
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Amount");
        sheet.Cells["A2"].PutValue("Food");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["A3"].PutValue("Drink");
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["A4"].PutValue("Food");
        sheet.Cells["B4"].PutValue(150);
        sheet.Cells["A5"].PutValue("Drink");
        sheet.Cells["B5"].PutValue(70);

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

        // Disable UI elements related to the pivot table
        pivotTable.EnableWizard = false;        // Hide the PivotTable wizard ribbon
        pivotTable.EnableFieldList = false;     // Hide the field list toolbar
        workbook.Settings.HidePivotFieldList = true; // Globally hide the pivot field list

        // Save the workbook as ODS with the minimal UI settings
        workbook.Save("PivotTable_MinimalUI.ods", SaveFormat.Ods);
    }
}