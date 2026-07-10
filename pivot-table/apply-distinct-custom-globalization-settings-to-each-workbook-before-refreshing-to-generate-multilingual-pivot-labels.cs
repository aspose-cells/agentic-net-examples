using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsPivotGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // First workbook – English-like labels
            Workbook wbEn = new Workbook();
            Worksheet wsEn = wbEn.Worksheets[0];
            PopulateSampleData(wsEn);
            CreatePivotTable(wsEn, "E5", "PivotEn");

            // Configure English-like globalization settings
            SettableGlobalizationSettings enSettings = new SettableGlobalizationSettings();
            SettablePivotGlobalizationSettings enPivotSettings = new SettablePivotGlobalizationSettings();
            enPivotSettings.SetTextOfColumnLabels("Column Headers");
            enPivotSettings.SetTextOfRowLabels("Row Headers");
            enPivotSettings.SetTextOfTotal("Total Amount");
            enPivotSettings.SetTextOfGrandTotal("Grand Total Amount");
            enPivotSettings.SetTextOfMultipleItems("Multiple Items Selected");
            enSettings.PivotSettings = enPivotSettings;
            wbEn.Settings.GlobalizationSettings = enSettings;

            // Refresh pivot to apply settings
            wsEn.PivotTables[0].RefreshData();
            wsEn.PivotTables[0].CalculateData();

            // Save the first workbook
            wbEn.Save("Pivot_English.xlsx");

            // Second workbook – French-like labels
            Workbook wbFr = new Workbook();
            Worksheet wsFr = wbFr.Worksheets[0];
            PopulateSampleData(wsFr);
            CreatePivotTable(wsFr, "E5", "PivotFr");

            // Configure French-like globalization settings
            SettableGlobalizationSettings frSettings = new SettableGlobalizationSettings();
            SettablePivotGlobalizationSettings frPivotSettings = new SettablePivotGlobalizationSettings();
            frPivotSettings.SetTextOfColumnLabels("En-têtes de colonne");
            frPivotSettings.SetTextOfRowLabels("En-têtes de ligne");
            frPivotSettings.SetTextOfTotal("Montant total");
            frPivotSettings.SetTextOfGrandTotal("Total général");
            frPivotSettings.SetTextOfMultipleItems("Éléments multiples");
            frSettings.PivotSettings = frPivotSettings;
            wbFr.Settings.GlobalizationSettings = frSettings;

            // Refresh pivot to apply French settings
            wsFr.PivotTables[0].RefreshData();
            wsFr.PivotTables[0].CalculateData();

            // Save the second workbook
            wbFr.Save("Pivot_French.xlsx");
        }

        // Helper method to add sample data to a worksheet
        private static void PopulateSampleData(Worksheet sheet)
        {
            // Header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            // Data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Apple");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(1500);

            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(800);

            sheet.Cells["A5"].PutValue("Orange");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(950);
        }

        // Helper method to create a simple pivot table
        private static void CreatePivotTable(Worksheet sheet, string destinationCell, string pivotName)
        {
            // Define source range (including headers)
            string sourceRange = "A1:C5";

            // Add pivot table
            int pivotIndex = sheet.PivotTables.Add(sourceRange, destinationCell, pivotName);
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields: Product as row, Region as column, Sales as data
            pivot.AddFieldToArea(PivotFieldType.Row, 0);      // Product
            pivot.AddFieldToArea(PivotFieldType.Column, 1);   // Region
            pivot.AddFieldToArea(PivotFieldType.Data, 2);     // Sales
        }
    }
}