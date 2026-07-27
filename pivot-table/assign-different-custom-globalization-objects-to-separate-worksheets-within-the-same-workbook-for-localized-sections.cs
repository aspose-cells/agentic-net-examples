using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook with three worksheets
        Workbook workbook = new Workbook();
        Worksheet sheetEn = workbook.Worksheets[0];
        sheetEn.Name = "English";

        Worksheet sheetRu = workbook.Worksheets.Add("Russian");
        Worksheet sheetDe = workbook.Worksheets.Add("German");

        // ---------- Russian worksheet ----------
        // Create a custom globalization object for Russian
        SettableGlobalizationSettings ruSettings = new SettableGlobalizationSettings();
        ruSettings.SetBooleanValueString(true, "ИСТИНА");
        ruSettings.SetBooleanValueString(false, "ЛОЖЬ");
        ruSettings.SetLocalFunctionName("SUM", "СУММ", true); // map SUM to СУММ

        // Apply Russian settings and fill the sheet
        workbook.Settings.GlobalizationSettings = ruSettings;
        FillSheet(sheetRu, "RU");

        // ---------- German worksheet ----------
        // Create a custom globalization object for German
        SettableGlobalizationSettings deSettings = new SettableGlobalizationSettings();
        deSettings.SetBooleanValueString(true, "WAHR");
        deSettings.SetBooleanValueString(false, "FALSCH");
        deSettings.SetLocalFunctionName("SUM", "SUMME", true); // map SUM to SUMME

        // Apply German settings and fill the sheet
        workbook.Settings.GlobalizationSettings = deSettings;
        FillSheet(sheetDe, "DE");

        // ---------- English (default) worksheet ----------
        // Reset to default globalization settings
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings();
        FillSheet(sheetEn, "EN");

        // Save the workbook
        workbook.Save("LocalizedWorkbook.xlsx");
    }

    // Helper method to populate a worksheet with sample data
    static void FillSheet(Worksheet sheet, string prefix)
    {
        Cells cells = sheet.Cells;

        // Boolean values – will be displayed according to the active globalization settings
        cells[0, 0].PutValue(true);
        cells[0, 1].PutValue(false);

        // Numeric values for SUM demonstration
        cells[1, 0].PutValue(10);
        cells[2, 0].PutValue(20);
        cells[3, 0].PutValue(30);

        // Formula using the standard function name "SUM"
        // The active globalization settings map it to the localized name when needed
        cells[4, 0].Formula = "=SUM(A2:A4)";

        // Ensure formulas are calculated with the current settings
        sheet.Workbook.CalculateFormula();
    }
}