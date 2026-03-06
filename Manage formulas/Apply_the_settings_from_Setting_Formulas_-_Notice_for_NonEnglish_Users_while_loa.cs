using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create load options for an XLSX file
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Set regional settings for a non‑English locale (example: Russian)
        loadOptions.CultureInfo = new CultureInfo("ru-RU");
        loadOptions.Region = CountryCode.Russia;
        loadOptions.LanguageCode = CountryCode.Russia;

        // Skip automatic formula parsing on load (optional, improves performance)
        loadOptions.ParsingFormulaOnOpen = false;

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Apply custom globalization settings to handle non‑English formula texts
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Ensure formulas are not recalculated automatically when the file is opened
        workbook.Settings.FormulaSettings.CalculateOnOpen = false;

        // Save the workbook after applying the settings
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }

    // Custom globalization settings matching the "Setting Formulas - Notice for Non‑English Users" example
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?": return "#ИМЯ?";
                case "#DIV/0!": return "#ДЕЛ/0!";
                case "#REF!": return "#ССЫЛКА!";
                case "#VALUE!": return "#ЗНАЧ!";
                case "#N/A": return "#Н/Д";
                case "#NUM!": return "#ЧИСЛО!";
                case "#NULL!": return "#ПУСТО!";
                default: return base.GetErrorValueString(err);
            }
        }
    }
}