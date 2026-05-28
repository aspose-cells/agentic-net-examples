using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPivotGrandTotalCustomization
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for pivot tables
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Region");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Electronics");
            dataSheet.Cells["B2"].PutValue("North");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("Electronics");
            dataSheet.Cells["B3"].PutValue("South");
            dataSheet.Cells["C3"].PutValue(1500);

            dataSheet.Cells["A4"].PutValue("Furniture");
            dataSheet.Cells["B4"].PutValue("North");
            dataSheet.Cells["C4"].PutValue(800);

            dataSheet.Cells["A5"].PutValue("Furniture");
            dataSheet.Cells["B5"].PutValue("South");
            dataSheet.Cells["C5"].PutValue(950);

            // Add first pivot table
            Worksheet pivotSheet1 = workbook.Worksheets.Add("Pivot1");
            int pivotIndex1 = pivotSheet1.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable1 = pivotSheet1.PivotTables[pivotIndex1];
            pivotTable1.AddFieldToArea(PivotFieldType.Row, 0);      // Category
            pivotTable1.AddFieldToArea(PivotFieldType.Column, 1);   // Region
            pivotTable1.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // Add second pivot table (different layout)
            Worksheet pivotSheet2 = workbook.Worksheets.Add("Pivot2");
            int pivotIndex2 = pivotSheet2.PivotTables.Add("A1:C5", "E3", "PivotTable2");
            PivotTable pivotTable2 = pivotSheet2.PivotTables[pivotIndex2];
            pivotTable2.AddFieldToArea(PivotFieldType.Row, 1);      // Region
            pivotTable2.AddFieldToArea(PivotFieldType.Column, 0);   // Category
            pivotTable2.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // Create custom globalization settings
            GlobalizationSettings globalization = new GlobalizationSettings();

            // Use SettablePivotGlobalizationSettings to change the Grand Total label
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();
            pivotSettings.SetTextOfGrandTotal("My Custom Grand Total");

            // Assign the customized pivot settings to the workbook's globalization settings
            globalization.PivotSettings = pivotSettings;
            workbook.Settings.GlobalizationSettings = globalization;

            // Refresh and calculate all pivot tables to apply the new label
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (PivotTable pt in ws.PivotTables)
                {
                    pt.RefreshData();
                    pt.CalculateData();
                }
            }

            // Save the workbook
            workbook.Save("PivotGrandTotalCustom.xlsx");
        }
    }
}