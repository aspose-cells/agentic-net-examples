using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsMultilingualPivot
{
    class Program
    {
        static void Main()
        {
            // Create workbooks with different language settings
            CreateWorkbookWithCustomPivotLabels("Pivot_EN.xlsx", Language.English);
            CreateWorkbookWithCustomPivotLabels("Pivot_FR.xlsx", Language.French);
            CreateWorkbookWithCustomPivotLabels("Pivot_DE.xlsx", Language.German);
        }

        // Enum to identify language variants
        enum Language
        {
            English,
            French,
            German
        }

        // Creates a workbook, applies language‑specific pivot globalization settings,
        // builds a simple pivot table, refreshes it and saves the file.
        static void CreateWorkbookWithCustomPivotLabels(string fileName, Language lang)
        {
            // ----- Workbook and sample data -----
            Workbook wb = new Workbook();
            Worksheet dataSheet = wb.Worksheets[0];
            Cells cells = dataSheet.Cells;

            // Sample data: Category | Amount
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("Fruit");
            cells["B2"].PutValue(1200);
            cells["A3"].PutValue("Vegetable");
            cells["B3"].PutValue(800);
            cells["A4"].PutValue("Fruit");
            cells["B4"].PutValue(1500);
            cells["A5"].PutValue("Vegetable");
            cells["B5"].PutValue(950);

            // ----- Globalization settings -----
            // General settings container
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

            // Pivot‑specific settings container
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // Apply language‑specific texts
            switch (lang)
            {
                case Language.English:
                    pivotSettings.SetTextOfColumnLabels("Column Labels");
                    pivotSettings.SetTextOfRowLabels("Row Labels");
                    pivotSettings.SetTextOfTotal("Total");
                    pivotSettings.SetTextOfGrandTotal("Grand Total");
                    pivotSettings.SetTextOfMultipleItems("(Multiple Items)");
                    break;

                case Language.French:
                    pivotSettings.SetTextOfColumnLabels("Étiquettes de colonne");
                    pivotSettings.SetTextOfRowLabels("Étiquettes de ligne");
                    pivotSettings.SetTextOfTotal("Total");
                    pivotSettings.SetTextOfGrandTotal("Total général");
                    pivotSettings.SetTextOfMultipleItems("(Éléments multiples)");
                    break;

                case Language.German:
                    pivotSettings.SetTextOfColumnLabels("Spaltenbeschriftungen");
                    pivotSettings.SetTextOfRowLabels("Zeilenbeschriftungen");
                    pivotSettings.SetTextOfTotal("Summe");
                    pivotSettings.SetTextOfGrandTotal("Gesamtsumme");
                    pivotSettings.SetTextOfMultipleItems("(Mehrere Elemente)");
                    break;
            }

            // Attach pivot settings to the global settings
            globalSettings.PivotSettings = pivotSettings;

            // Assign the globalization settings to the workbook
            wb.Settings.GlobalizationSettings = globalSettings;

            // ----- Pivot table creation -----
            Worksheet pivotSheet = wb.Worksheets.Add("PivotTable");
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B5", "A1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: Category as row, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Amount

            // Refresh and calculate to apply the custom labels
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ----- Save the workbook -----
            wb.Save(fileName);
        }
    }
}