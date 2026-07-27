using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Q1");
        sheet.Cells["C1"].PutValue("Q2");
        sheet.Cells["A2"].PutValue("Electronics");
        sheet.Cells["B2"].PutValue(1500);
        sheet.Cells["C2"].PutValue(2200);
        sheet.Cells["A3"].PutValue("Clothing");
        sheet.Cells["B3"].PutValue(900);
        sheet.Cells["C3"].PutValue(1300);

        // Add a pivot table based on the sample data
        int pivotIndex = sheet.PivotTables.Add("A1:C3", "E4", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Q1
        pivot.AddFieldToArea(PivotFieldType.Data, 2);  // Q2

        // Create custom globalization settings to change the Grand Total label
        SettablePivotGlobalizationSettings customSettings = new SettablePivotGlobalizationSettings();
        customSettings.SetTextOfGrandTotal("My Custom Grand Total");

        // Assign the custom settings to the workbook's globalization settings
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings();
        workbook.Settings.GlobalizationSettings.PivotSettings = customSettings;

        // Refresh and calculate all pivot tables in the workbook to apply the new label
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (PivotTable pt in ws.PivotTables)
            {
                pt.RefreshData();
                pt.CalculateData();
            }
        }

        // Save the workbook
        workbook.Save("CustomGrandTotal.xlsx");
    }
}