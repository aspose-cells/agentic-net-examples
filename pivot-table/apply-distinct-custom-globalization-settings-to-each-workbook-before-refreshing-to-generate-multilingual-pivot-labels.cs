// Title: How to set different SettableGlobalizationSettings for English and French pivot tables using Aspose.Cells in C#
// AI Prompts: Generate two workbooks with identical sales data, add a pivot table to each, and apply distinct SettableGlobalizationSettings for English and French labels before refreshing and saving the files. | Create a helper method that takes custom label strings and returns a configured SettablePivotGlobalizationSettings object, then use it to assign language‑specific globalization settings to a workbook's pivot tables. | Extend the sample to produce a third workbook with Spanish pivot labels by reusing the helper method and saving the result as a separate Excel file.
// Common Searches: Aspose.Cells C# set custom pivot table labels per workbook | How to apply different globalization settings to multiple Excel files using Aspose.Cells | C# example for multilingual pivot tables with SettablePivotGlobalizationSettings | Refresh pivot after changing SettableGlobalizationSettings in Aspose.Cells | Create English and French pivot tables in the same Aspose.Cells project
// Tags: settableglobalizationsettings for pivot tables | aspocells multilingual pivot labels | c# apply language specific pivot globalization | refresh pivot after globalization change | excel workbook separate language settings

using System;
using Aspose.Cells;
using Aspose.Cells.Settings;
using Aspose.Cells.Pivot;

namespace AsposeCellsMultilingualPivot
{
    // The example creates two workbooks with identical sales data, adds a pivot table to each, configures separate SettableGlobalizationSettings for English and French (customizing labels such as Total, Grand Total, Row Labels, etc.), refreshes and calculates the pivots, and saves the workbooks as distinct Excel files.
    class Program
    {
        static void Main()
        {
            // ==============================
            // Workbook 1 – English labels
            // ==============================
            Workbook wbEn = new Workbook();
            Worksheet wsEn = wbEn.Worksheets[0];

            // Sample data
            wsEn.Cells["A1"].PutValue("Product");
            wsEn.Cells["B1"].PutValue("Region");
            wsEn.Cells["C1"].PutValue("Sales");
            wsEn.Cells["A2"].PutValue("Apple");
            wsEn.Cells["B2"].PutValue("North");
            wsEn.Cells["C2"].PutValue(1200);
            wsEn.Cells["A3"].PutValue("Apple");
            wsEn.Cells["B3"].PutValue("South");
            wsEn.Cells["C3"].PutValue(800);
            wsEn.Cells["A4"].PutValue("Orange");
            wsEn.Cells["B4"].PutValue("North");
            wsEn.Cells["C4"].PutValue(1500);
            wsEn.Cells["A5"].PutValue("Orange");
            wsEn.Cells["B5"].PutValue("South");
            wsEn.Cells["C5"].PutValue(900);

            // Create pivot table
            int pivotIdxEn = wsEn.PivotTables.Add("A1:C5", "E2", "SalesPivot_EN");
            PivotTable pivotEn = wsEn.PivotTables[pivotIdxEn];
            pivotEn.AddFieldToArea(PivotFieldType.Row, 0);      // Product
            pivotEn.AddFieldToArea(PivotFieldType.Column, 1);   // Region
            pivotEn.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // Configure English globalization settings
            SettableGlobalizationSettings gSettingsEn = new SettableGlobalizationSettings();
            SettablePivotGlobalizationSettings pSettingsEn = new SettablePivotGlobalizationSettings();

            pSettingsEn.SetTextOfTotal("Total");
            pSettingsEn.SetTextOfGrandTotal("Grand Total");
            pSettingsEn.SetTextOfRowLabels("Row Labels");
            pSettingsEn.SetTextOfColumnLabels("Column Labels");
            pSettingsEn.SetTextOfMultipleItems("(Multiple Items)");
            pSettingsEn.SetTextOfAll("All");

            gSettingsEn.PivotSettings = pSettingsEn;
            wbEn.Settings.GlobalizationSettings = gSettingsEn;

            // Refresh and calculate pivot to apply settings
            pivotEn.RefreshData();
            pivotEn.CalculateData();

            // Save English workbook
            wbEn.Save("Pivot_Multilingual_EN.xlsx");

            // ==============================
            // Workbook 2 – French labels
            // ==============================
            Workbook wbFr = new Workbook();
            Worksheet wsFr = wbFr.Worksheets[0];

            // Same sample data
            wsFr.Cells["A1"].PutValue("Produit");
            wsFr.Cells["B1"].PutValue("Région");
            wsFr.Cells["C1"].PutValue("Ventes");
            wsFr.Cells["A2"].PutValue("Pomme");
            wsFr.Cells["B2"].PutValue("Nord");
            wsFr.Cells["C2"].PutValue(1200);
            wsFr.Cells["A3"].PutValue("Pomme");
            wsFr.Cells["B3"].PutValue("Sud");
            wsFr.Cells["C3"].PutValue(800);
            wsFr.Cells["A4"].PutValue("Orange");
            wsFr.Cells["B4"].PutValue("Nord");
            wsFr.Cells["C4"].PutValue(1500);
            wsFr.Cells["A5"].PutValue("Orange");
            wsFr.Cells["B5"].PutValue("Sud");
            wsFr.Cells["C5"].PutValue(900);

            // Create pivot table
            int pivotIdxFr = wsFr.PivotTables.Add("A1:C5", "E2", "SalesPivot_FR");
            PivotTable pivotFr = wsFr.PivotTables[pivotIdxFr];
            pivotFr.AddFieldToArea(PivotFieldType.Row, 0);      // Produit
            pivotFr.AddFieldToArea(PivotFieldType.Column, 1);   // Région
            pivotFr.AddFieldToArea(PivotFieldType.Data, 2);     // Ventes

            // Configure French globalization settings
            SettableGlobalizationSettings gSettingsFr = new SettableGlobalizationSettings();
            SettablePivotGlobalizationSettings pSettingsFr = new SettablePivotGlobalizationSettings();

            pSettingsFr.SetTextOfTotal("Total");
            pSettingsFr.SetTextOfGrandTotal("Total Général");
            pSettingsFr.SetTextOfRowLabels("Étiquettes de lignes");
            pSettingsFr.SetTextOfColumnLabels("Étiquettes de colonnes");
            pSettingsFr.SetTextOfMultipleItems("(Éléments multiples)");
            pSettingsFr.SetTextOfAll("Tous");

            gSettingsFr.PivotSettings = pSettingsFr;
            wbFr.Settings.GlobalizationSettings = gSettingsFr;

            // Refresh and calculate pivot to apply French labels
            pivotFr.RefreshData();
            pivotFr.CalculateData();

            // Save French workbook
            wbFr.Save("Pivot_Multilingual_FR.xlsx");
        }
    }
}
